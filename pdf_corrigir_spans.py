"""
pdf_corrigir_spans.py
---------------------
Detecta e corrige automaticamente linhas com spans justificados separados
(ex: 'YARA ... SOUZA' em 4 spans espalhados) consolidando-as num unico
bloco de texto contiguamente posicionado.

Uso:
    python pdf_corrigir_spans.py                  # abre dialogo de arquivo
    python pdf_corrigir_spans.py arquivo.pdf      # processa diretamente
    python pdf_corrigir_spans.py pasta/           # processa todos os PDFs da pasta
"""
from __future__ import annotations

import os
import sys
import re
from pathlib import Path
from collections import defaultdict
from dataclasses import dataclass

import fitz


# ---------------------------------------------------------------------------
# Estruturas de dados
# ---------------------------------------------------------------------------

@dataclass
class FragmentedLine:
    """Linha visual com multiplos spans espalhados que precisam ser consolidados."""
    page_num:   int
    bbox:       tuple[float, float, float, float]   # bbox total da linha
    text:       str                                  # texto completo
    font_name:  str
    font_size:  float
    font_flags: int
    color_int:  int
    raw_spans:  list[dict]


# ---------------------------------------------------------------------------
# Deteccao
# ---------------------------------------------------------------------------

def _e_nome_justificado(gaps: list[float], block_bbox: tuple, page_width: float,
                        linhas_visuais_no_bloco: int, spans: list[dict],
                        outras_linhas_tem_multiplos_spans: bool = False) -> bool:
    """
    Retorna True somente se a linha parecer um nome pessoal justificado
    (nao uma tabela de horarios, cabecalho ou linha de filhos).

    Criterios:
    1. O bloco nao contem outras linhas visuais com spans espalhados — nomes de
       assinatura podem compartilhar bloco com linhas simples (ex: '________'),
       mas tabelas de horario tem multiplas linhas todas com spans espalhados.
    2. Gaps uniformes entre TODOS os spans (variacao < 15pt).
    3. O bloco ocupa > 55% da largura da pagina.
    4. Nao e apenas 2 spans com gap unico > 150pt (coluna de tabela).
    5. O ultimo span nao parece uma data (DD/MM/AAAA).
    """
    import re as _re

    if len(gaps) < 1:
        return False

    # Criterio 1: outras linhas do bloco nao podem ter spans espalhados
    # (indica tabela de horarios com multiplas linhas fragmentadas)
    if outras_linhas_tem_multiplos_spans:
        return False

    # Criterio 4: rejeita par de colunas com gap gigante
    if len(gaps) == 1 and gaps[0] > 150:
        return False

    # Criterio 5: rejeita se o ultimo span parece uma data
    ultimo_texto = spans[-1].get("text", "").strip()
    if _re.match(r"\d{2}/\d{2}/\d{4}", ultimo_texto):
        return False

    # Criterio 2: gaps uniformes — compara todos os gaps
    # Se o ultimo gap e muito maior que a media dos anteriores, e nome+data
    if len(gaps) >= 2:
        gaps_sem_ultimo = gaps[:-1]
        media_anterior  = sum(gaps_sem_ultimo) / len(gaps_sem_ultimo)
        if gaps[-1] > media_anterior * 3:
            return False

    gap_variacao   = max(gaps) - min(gaps)
    gaps_uniformes = gap_variacao < 15.0

    # Criterio 3: bloco de largura quase total da pagina
    block_width = block_bbox[2] - block_bbox[0]
    bloco_largo = (block_width / page_width) > 0.55

    return gaps_uniformes and bloco_largo


def detectar_linhas_fragmentadas(doc: fitz.Document, min_spans: int = 2) -> list[FragmentedLine]:
    """
    Varre o documento e retorna linhas de nomes justificados fragmentados.
    Tabelas e cabecalhos com colunas sao ignorados automaticamente.
    """
    resultado: list[FragmentedLine] = []

    for page_num in range(len(doc)):
        page       = doc[page_num]
        page_width = page.rect.width
        raw        = page.get_text("dict")

        for block in raw.get("blocks", []):
            if block.get("type") != 0:
                continue

            block_bbox = tuple(block["bbox"])

            # Agrupa linhas do bloco por y_center (tolerancia 2pt)
            row_map: dict[float, list[dict]] = {}
            for line in block.get("lines", []):
                y0, y1 = line["bbox"][1], line["bbox"][3]
                y_key  = round((y0 + y1) / 2, 0)
                matched = next((k for k in row_map if abs(k - y_key) <= 2), None)
                if matched is None:
                    matched = y_key
                    row_map[matched] = []
                row_map[matched].append(line)

            for y_key in sorted(row_map):
                row_lines = row_map[y_key]
                spans_with_text = [
                    s for line in row_lines
                    for s in line.get("spans", [])
                    if s.get("text", "").strip()
                ]

                if len(spans_with_text) < min_spans:
                    continue

                spans_with_text.sort(key=lambda s: s["bbox"][0])
                gaps = [
                    spans_with_text[i+1]["bbox"][0] - spans_with_text[i]["bbox"][2]
                    for i in range(len(spans_with_text) - 1)
                ]

                # So processa se todos os gaps forem > 10pt
                if not all(g > 10 for g in gaps):
                    continue

                # Filtra: so nomes justificados, nao tabelas
                linhas_visuais_no_bloco = len(row_map)

                # Verifica se outras linhas visuais do bloco tambem tem spans espalhados
                # (sinal de tabela, nao de nome compartilhando bloco com linha de tracos)
                outras_com_multiplos = False
                for outro_y, outras_lines in row_map.items():
                    if abs(outro_y - y_key) < 2:
                        continue  # e a propria linha atual
                    outros_spans = [
                        s for l in outras_lines
                        for s in l.get("spans", [])
                        if s.get("text", "").strip()
                    ]
                    if len(outros_spans) >= 2:
                        outros_spans.sort(key=lambda s: s["bbox"][0])
                        outros_gaps = [
                            outros_spans[i+1]["bbox"][0] - outros_spans[i]["bbox"][2]
                            for i in range(len(outros_spans) - 1)
                        ]
                        if any(g > 10 for g in outros_gaps):
                            outras_com_multiplos = True
                            break

                if not _e_nome_justificado(gaps, block_bbox, page_width,
                                           linhas_visuais_no_bloco, spans_with_text,
                                           outras_com_multiplos):
                    continue

                joined = " ".join(s["text"].strip() for s in spans_with_text)
                text   = re.sub(r" {2,}", " ", joined).strip()
                if not text:
                    continue

                xs0 = [s["bbox"][0] for s in spans_with_text]
                ys0 = [s["bbox"][1] for s in spans_with_text]
                xs1 = [s["bbox"][2] for s in spans_with_text]
                ys1 = [s["bbox"][3] for s in spans_with_text]

                first = spans_with_text[0]
                resultado.append(FragmentedLine(
                    page_num   = page_num,
                    bbox       = (min(xs0), min(ys0), max(xs1), max(ys1)),
                    text       = text,
                    font_name  = first.get("font", "helv"),
                    font_size  = first.get("size", 11.0),
                    font_flags = first.get("flags", 0),
                    color_int  = first.get("color", 0),
                    raw_spans  = spans_with_text,
                ))

    return resultado


def detectar_nomes_filhos(doc: fitz.Document) -> list[FragmentedLine]:
    """
    Detecta linhas da tabela de dependentes onde o nome do filho esta
    fragmentado em spans separados com uma data no final.

    Padrao: ['RHAVI','FELIPE','LEMOS','PAULO','DA','SILVA','04/02/2024']
    - Todos os gaps entre as palavras do nome sao uniformes e pequenos (< 60pt)
    - O ultimo span e uma data DD/MM/AAAA
    - Ha um gap notavelmente maior antes da data

    Retorna FragmentedLine apenas para os spans do NOME (sem a data),
    com bbox cobrindo so a area do nome (nao inclui a data).
    """
    resultado: list[FragmentedLine] = []
    data_pat = re.compile(r"^\d{2}/\d{2}/\d{4}$")

    for page_num in range(len(doc)):
        page = doc[page_num]
        raw  = page.get_text("dict")

        for block in raw.get("blocks", []):
            if block.get("type") != 0:
                continue

            row_map: dict[float, list[dict]] = {}
            for line in block.get("lines", []):
                y0, y1 = line["bbox"][1], line["bbox"][3]
                y_key  = round((y0 + y1) / 2, 0)
                matched = next((k for k in row_map if abs(k - y_key) <= 2), None)
                if matched is None:
                    matched = y_key
                    row_map[matched] = []
                row_map[matched].append(line)

            for y_key in sorted(row_map):
                spans = [
                    s for l in row_map[y_key]
                    for s in l.get("spans", [])
                    if s.get("text", "").strip()
                ]
                if len(spans) < 3:
                    continue

                spans.sort(key=lambda s: s["bbox"][0])

                # Ultimo span deve ser data
                if not data_pat.match(spans[-1]["text"].strip()):
                    continue

                # Spans do nome = todos exceto o ultimo (data)
                nome_spans = spans[:-1]
                if len(nome_spans) < 2:
                    continue

                gaps_nome = [
                    nome_spans[i+1]["bbox"][0] - nome_spans[i]["bbox"][2]
                    for i in range(len(nome_spans) - 1)
                ]

                # Gaps do nome devem ser uniformes e pequenos (< 60pt)
                if not all(g > 5 for g in gaps_nome):
                    continue
                if max(gaps_nome) > 60:
                    continue
                if (max(gaps_nome) - min(gaps_nome)) > 15:
                    continue

                # Gap antes da data deve ser maior que os gaps do nome
                gap_antes_data = spans[-1]["bbox"][0] - nome_spans[-1]["bbox"][2]
                if gap_antes_data <= max(gaps_nome) * 2:
                    continue

                joined = " ".join(s["text"].strip() for s in nome_spans)
                text   = re.sub(r" {2,}", " ", joined).strip()
                if not text:
                    continue

                xs0 = [s["bbox"][0] for s in nome_spans]
                ys0 = [s["bbox"][1] for s in nome_spans]
                xs1 = [s["bbox"][2] for s in nome_spans]
                ys1 = [s["bbox"][3] for s in nome_spans]

                first = nome_spans[0]
                resultado.append(FragmentedLine(
                    page_num   = page_num,
                    bbox       = (min(xs0), min(ys0), max(xs1), max(ys1)),
                    text       = text,
                    font_name  = first.get("font", "helv"),
                    font_size  = first.get("size", 11.0),
                    font_flags = first.get("flags", 0),
                    color_int  = first.get("color", 0),
                    raw_spans  = nome_spans,
                ))

    return resultado


# ---------------------------------------------------------------------------
# Correcao
# ---------------------------------------------------------------------------

def _int_to_rgb(color_int: int) -> tuple[float, float, float]:
    r = ((color_int >> 16) & 0xFF) / 255.0
    g = ((color_int >>  8) & 0xFF) / 255.0
    b = (color_int         & 0xFF) / 255.0
    return (r, g, b)


def _map_font(font_name: str, flags: int) -> str:
    is_bold   = bool(flags & 16)
    is_italic = bool(flags & 2)
    name = font_name.lower()
    if "+" in name:
        name = name.split("+", 1)[1]
    name = name.replace(",bold","").replace(",italic","").replace(",bolditalic","").strip()
    mono  = ("courier","consolas","monaco","monospace")
    sans  = ("arial","helvetica","calibri","verdana","tahoma","trebuchet","gill")
    if any(n in name for n in mono):
        return ("cobi","cobo","coit","cour")[is_bold*2 + is_italic*1 if not (is_bold and is_italic) else 0]
    if any(n in name for n in sans):
        if is_bold and is_italic: return "hebi"
        if is_bold:               return "hebo"
        if is_italic:             return "heit"
        return "helv"
    if is_bold and is_italic: return "tibi"
    if is_bold:               return "tibo"
    if is_italic:             return "tiro"
    return "tiro"


def consolidar_linha(doc: fitz.Document, linha: FragmentedLine) -> str | None:
    """
    Apaga os spans fragmentados e reescreve o texto consolidado.
    Usa insert_text (posicional por baseline) para evitar restricao de altura
    do insert_textbox em linhas com bbox muito estreito (< 1 linha de altura).
    Retorna mensagem de aviso em caso de erro, None se OK.
    """
    page = doc[linha.page_num]

    # Redact individual em cada span (preserva o restante da pagina)
    for span in linha.raw_spans:
        page.add_redact_annot(fitz.Rect(span["bbox"]), fill=(1, 1, 1))
    page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)

    # Baseline = y1 do bbox (base inferior do texto)
    x0, y0, x1, y1 = linha.bbox
    baseline = fitz.Point(x0, y1 - 1)   # 1pt acima do fundo do bbox

    color    = _int_to_rgb(linha.color_int)
    fontname = _map_font(linha.font_name, linha.font_flags)

    rc = page.insert_text(
        baseline,
        linha.text,
        fontname = fontname,
        fontsize = linha.font_size,
        color    = color,
        overlay  = True,
    )
    if rc < 1:
        return f"Pag {linha.page_num+1}: '{linha.text[:30]}' — erro ao inserir (rc={rc})"
    return None


def corrigir_pdf(input_path: Path, output_path: Path | None = None,
                 output_dir: Path | None = None) -> dict:
    """
    Detecta e corrige todas as linhas fragmentadas de um PDF.
    Retorna dict com estatisticas: {'fragmentadas': N, 'corrigidas': N, 'avisos': [...]}

    output_dir: se fornecido, salva o arquivo corrigido dentro dessa pasta
                (com o mesmo nome do original, sem sufixo).
    output_path: caminho completo de saida (tem precedencia sobre output_dir).
    """
    if output_path is None:
        if output_dir is not None:
            output_dir.mkdir(parents=True, exist_ok=True)
            output_path = output_dir / input_path.name
        else:
            output_path = input_path.with_stem(input_path.stem + "_corrigido")

    doc    = fitz.open(str(input_path))
    linhas = detectar_linhas_fragmentadas(doc) + detectar_nomes_filhos(doc)
    avisos     = []
    corrigidas = 0

    if not linhas:
        doc.close()
        return {"fragmentadas": 0, "corrigidas": 0, "avisos": [], "output": None}

    # Processa de baixo pra cima por pagina (evita deslocamento de coordenadas)
    by_page: dict[int, list[FragmentedLine]] = defaultdict(list)
    for l in linhas:
        by_page[l.page_num].append(l)

    for page_num in sorted(by_page):
        linhas_pag = sorted(by_page[page_num], key=lambda l: l.bbox[3], reverse=True)
        for linha in linhas_pag:
            aviso = consolidar_linha(doc, linha)
            if aviso:
                avisos.append(aviso)
            corrigidas += 1

    # PyMuPDF nao permite doc.save() de volta no mesmo caminho ja aberto
    # com garbage/clean/deflate (exige incremental=True nesse caso). Para
    # poder sobrescrever o original com a otimizacao completa, salva-se
    # num arquivo temporario na mesma pasta e substitui atomicamente.
    if output_path.resolve() == input_path.resolve():
        tmp_path = output_path.with_name(output_path.stem + ".tmp_corrigido" + output_path.suffix)
        doc.save(str(tmp_path), garbage=4, deflate=True, clean=True)
        doc.close()
        os.replace(tmp_path, output_path)
    else:
        doc.save(str(output_path), garbage=4, deflate=True, clean=True)
        doc.close()

    return {
        "fragmentadas": len(linhas),
        "corrigidas":   corrigidas,
        "avisos":       avisos,
        "output":       output_path,
    }


# ---------------------------------------------------------------------------
# Relatorio / preview (sem modificar o PDF)
# ---------------------------------------------------------------------------

def relatorio_pdf(input_path: Path) -> None:
    """Mostra o que seria corrigido sem modificar o arquivo."""
    doc    = fitz.open(str(input_path))
    linhas = detectar_linhas_fragmentadas(doc) + detectar_nomes_filhos(doc)
    doc.close()

    print(f"\n{'='*60}")
    print(f"Arquivo: {input_path.name}")
    print(f"{'='*60}")

    if not linhas:
        print("  Nenhuma linha fragmentada detectada. PDF esta correto.")
        return

    print(f"  {len(linhas)} linha(s) fragmentada(s) detectada(s):\n")
    for l in linhas:
        print(f"  Pag {l.page_num+1} | {len(l.raw_spans)} spans | texto: {repr(l.text)}")
        gaps = [
            l.raw_spans[i+1]["bbox"][0] - l.raw_spans[i]["bbox"][2]
            for i in range(len(l.raw_spans)-1)
        ]
        print(f"         espacos entre spans: {[round(g,1) for g in gaps]}")


# ---------------------------------------------------------------------------
# Interface de linha de comando / UI simples com tkinter
# ---------------------------------------------------------------------------

def _processar_arquivo_ui(input_path: Path, log_fn, output_dir: Path | None = None) -> None:
    log_fn(f"\nProcessando: {input_path.name}")

    doc    = fitz.open(str(input_path))
    linhas = detectar_linhas_fragmentadas(doc) + detectar_nomes_filhos(doc)
    doc.close()

    if not linhas:
        log_fn(f"  OK — nenhuma fragmentacao encontrada.")
        return

    log_fn(f"  Encontradas {len(linhas)} linha(s) fragmentada(s):")
    for l in linhas:
        log_fn(f"    Pag {l.page_num+1}: {repr(l.text)}  ({len(l.raw_spans)} spans)")

    stats = corrigir_pdf(input_path, output_dir=output_dir)
    log_fn(f"  Corrigidas: {stats['corrigidas']}")
    if stats["avisos"]:
        for a in stats["avisos"]:
            log_fn(f"  AVISO: {a}")
    log_fn(f"  Salvo em: {stats['output']}")


def main_ui() -> None:
    import tkinter as tk
    from tkinter import ttk, filedialog, scrolledtext

    root = tk.Tk()
    root.title("Corretor de Spans PDF")
    root.geometry("720x520")

    # --- Toolbar ---
    tb = tk.Frame(root, bd=1, relief=tk.GROOVE)
    tb.pack(fill=tk.X, padx=8, pady=6)

    tk.Button(tb, text="Selecionar PDF(s)", command=lambda: _btn_selecionar(),
              width=18).pack(side=tk.LEFT, padx=4)
    tk.Button(tb, text="Selecionar Pasta", command=lambda: _btn_pasta(),
              width=18).pack(side=tk.LEFT, padx=4)
    tk.Button(tb, text="Apenas Relatorio", command=lambda: _btn_relatorio(),
              width=18).pack(side=tk.LEFT, padx=4)
    tk.Button(tb, text="Limpar", command=lambda: log_box.delete("1.0", tk.END),
              width=10).pack(side=tk.LEFT, padx=4)

    # --- Log ---
    log_box = scrolledtext.ScrolledText(root, wrap=tk.WORD, font=("Consolas", 9),
                                        state=tk.NORMAL)
    log_box.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

    # --- Status ---
    status_var = tk.StringVar(value="Selecione um ou mais PDFs para corrigir")
    tk.Label(root, textvariable=status_var, anchor="w",
             relief=tk.SUNKEN, padx=8, pady=4).pack(fill=tk.X, side=tk.BOTTOM)

    def log(msg: str) -> None:
        log_box.insert(tk.END, msg + "\n")
        log_box.see(tk.END)
        root.update_idletasks()

    def _btn_selecionar() -> None:
        paths = filedialog.askopenfilenames(
            title="Selecionar PDFs para corrigir",
            filetypes=[("PDF files", "*.pdf"), ("Todos", "*.*")],
        )
        if not paths:
            return
        # Pasta de saida: subpasta "corrigidos" ao lado do primeiro arquivo
        output_dir = Path(paths[0]).parent / "corrigidos"
        output_dir.mkdir(parents=True, exist_ok=True)
        log(f"\nPasta de saida: {output_dir}")
        status_var.set(f"Processando {len(paths)} arquivo(s)...")
        for p in paths:
            _processar_arquivo_ui(Path(p), log, output_dir=output_dir)
        status_var.set("Concluido.")

    def _btn_pasta() -> None:
        folder = filedialog.askdirectory(title="Selecionar pasta com PDFs")
        if not folder:
            return
        pdfs = list(Path(folder).glob("*.pdf"))
        if not pdfs:
            log("Nenhum PDF encontrado na pasta.")
            return
        # Pasta de saida: subpasta "corrigidos" dentro da pasta selecionada
        output_dir = Path(folder) / "corrigidos"
        output_dir.mkdir(parents=True, exist_ok=True)
        log(f"\nPasta de saida: {output_dir}")
        status_var.set(f"Processando {len(pdfs)} arquivo(s) da pasta...")
        for p in pdfs:
            _processar_arquivo_ui(p, log, output_dir=output_dir)
        status_var.set("Concluido.")

    def _btn_relatorio() -> None:
        paths = filedialog.askopenfilenames(
            title="Selecionar PDFs para analisar",
            filetypes=[("PDF files", "*.pdf"), ("Todos", "*.*")],
        )
        if not paths:
            return
        for p in paths:
            doc    = fitz.open(p)
            linhas = detectar_linhas_fragmentadas(doc)
            doc.close()
            name = Path(p).name
            log(f"\n{'='*50}")
            log(f"Relatorio: {name}")
            if not linhas:
                log("  Nenhuma fragmentacao detectada.")
            else:
                log(f"  {len(linhas)} linha(s) fragmentada(s):")
                for l in linhas:
                    log(f"    Pag {l.page_num+1}: {repr(l.text)}  ({len(l.raw_spans)} spans)")

    log("Corretor de Spans PDF")
    log("=" * 50)
    log("Detecta linhas com texto justificado fragmentado em")
    log("multiplos spans e consolida num unico bloco correto.")
    log("")
    log("- 'Selecionar PDF(s)': corrige e salva na pasta 'corrigidos'")
    log("- 'Selecionar Pasta': processa todos os PDFs, salva em 'corrigidos'")
    log("- 'Apenas Relatorio': mostra o que seria corrigido sem modificar")

    root.mainloop()


def main_cli() -> None:
    """Modo linha de comando."""
    args = sys.argv[1:]

    if not args:
        # Sem argumentos: abre UI
        main_ui()
        return

    for arg in args:
        path = Path(arg)
        if path.is_dir():
            pdfs = list(path.glob("*.pdf"))
            print(f"Pasta: {path} — {len(pdfs)} PDF(s)")
            for p in pdfs:
                relatorio_pdf(p)
                stats = corrigir_pdf(p)
                print(f"  Corrigidas: {stats['corrigidas']}  Salvo: {stats['output']}")
        elif path.is_file() and path.suffix.lower() == ".pdf":
            relatorio_pdf(path)
            stats = corrigir_pdf(path)
            print(f"\nCorrigidas: {stats['corrigidas']}")
            if stats["avisos"]:
                for a in stats["avisos"]:
                    print(f"AVISO: {a}")
            print(f"Salvo em: {stats['output']}")
        else:
            print(f"Ignorado (nao e PDF ou pasta): {arg}")


if __name__ == "__main__":
    main_cli()
