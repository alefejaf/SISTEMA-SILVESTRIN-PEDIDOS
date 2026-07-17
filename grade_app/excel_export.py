"""Geracao dos arquivos Excel: NOTA FINAL enxuta e relatorio completo."""
import io

import pandas as pd
from xlsxwriter.utility import xl_col_to_name

from .config import APP_NAME, preparar_config_grade


def gerar_excel_nota_final(nota_final: pd.DataFrame, base: pd.DataFrame, total_pedido_dia: float = 0.0, cliente_nome: str = "") -> bytes:
    """Gera um Excel limpo somente com a NOTA FINAL para digitação do pedido."""
    output = io.BytesIO()
    df = nota_final.copy()
    if "TOTAL" in df.columns:
        df = df[df["TOTAL"].fillna(0).astype(float).gt(0)].reset_index(drop=True)

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="NOTA_FINAL", index=False, startrow=2)
        workbook = writer.book
        ws = writer.sheets["NOTA_FINAL"]

        fmt_title = workbook.add_format({"bold": True, "font_size": 16, "font_color": "white", "bg_color": "#1F4E78", "align": "center", "valign": "vcenter"})
        fmt_header = workbook.add_format({"bold": True, "font_color": "white", "bg_color": "#1F4E78", "border": 1, "align": "center", "valign": "vcenter"})
        fmt_text = workbook.add_format({"border": 1, "valign": "vcenter"})
        fmt_qty = workbook.add_format({"num_format": "#,##0.00", "border": 1, "align": "center", "valign": "vcenter"})
        fmt_total = workbook.add_format({"num_format": "#,##0.00", "border": 1, "bold": True, "bg_color": "#D9EAD3", "align": "center"})
        fmt_money = workbook.add_format({"num_format": "R$ #,##0.00", "border": 1, "bold": True, "bg_color": "#FFF2CC", "align": "center"})

        last_col = max(len(df.columns) - 1, 0)
        titulo = f"NOTA FINAL - GRADE {cliente_nome}" if cliente_nome else "NOTA FINAL - GRADE"
        ws.merge_range(0, 0, 0, last_col, titulo, fmt_title)
        ws.write(1, 0, "Use esta aba para digitar/conferir o pedido: produto, quantidade por loja, total e preço.")

        # Localiza colunas de TOTAL e PREÇO para aplicar fórmulas
        total_col_idx = df.columns.get_loc("TOTAL") if "TOTAL" in df.columns else None
        preco_col_idx = df.columns.get_loc("PREÇO") if "PREÇO" in df.columns else None

        # Aplica a fórmula SOMA na coluna TOTAL para recálculos automáticos
        if total_col_idx is not None and len(df) > 0:
            last_store_col_idx = total_col_idx - 1
            start_letter = xl_col_to_name(1)  # coluna B
            end_letter = xl_col_to_name(last_store_col_idx)
            for i in range(len(df)):
                r_excel = i + 4  # dados começam na linha 4 física
                formula = f"=SUM({start_letter}{r_excel}:{end_letter}{r_excel})"
                ws.write_formula(i + 3, total_col_idx, formula, fmt_total)

        # TOTAL DO PEDIDO DO DIA via SUMPRODUCT
        if total_col_idx is not None and preco_col_idx is not None and len(df) > 0:
            total_letter = xl_col_to_name(total_col_idx)
            preco_letter = xl_col_to_name(preco_col_idx)
            ws.write(1, max(last_col - 1, 0), "TOTAL DO PEDIDO DO DIA")
            formula_total = f"=SUMPRODUCT({total_letter}4:{total_letter}{len(df) + 3}, {preco_letter}4:{preco_letter}{len(df) + 3})"
            ws.write_formula(1, last_col, formula_total, fmt_money)
        elif total_pedido_dia:
            ws.write(1, max(last_col - 1, 0), "TOTAL DO PEDIDO DO DIA")
            ws.write(1, last_col, total_pedido_dia, fmt_money)

        ws.freeze_panes(3, 1)
        ws.autofilter(2, 0, max(len(df) + 2, 3), last_col)
        ws.set_row(0, 26)
        ws.set_row(2, 24)

        for col_idx, col_name in enumerate(df.columns):
            ws.write(2, col_idx, col_name, fmt_header)
            if col_name == "Produto":
                ws.set_column(col_idx, col_idx, 34, fmt_text)
            elif col_name == "PREÇO":
                ws.set_column(col_idx, col_idx, 13, fmt_money)
            elif col_name == "TOTAL":
                ws.set_column(col_idx, col_idx, 12, fmt_total)
            else:
                ws.set_column(col_idx, col_idx, 10, fmt_qty)

        # Reaplica formatos por área para deixar a nota pronta para uso.
        if len(df) > 0:
            for row in range(3, len(df) + 3):
                ws.set_row(row, 20)
            produto_col = 0
            ws.set_column(produto_col, produto_col, 34, fmt_text)
            for col_idx, col_name in enumerate(df.columns):
                if col_name == "PREÇO":
                    ws.set_column(col_idx, col_idx, 13, fmt_money)
                elif col_name == "TOTAL":
                    ws.set_column(col_idx, col_idx, 12, fmt_total)
                elif col_name != "Produto":
                    ws.set_column(col_idx, col_idx, 10, fmt_qty)

        # Cria as abas de cada loja
        adicionar_abas_lojas(writer, base, workbook)
    return output.getvalue()

def gerar_excel(base: pd.DataFrame, grade: pd.DataFrame, validacoes: pd.DataFrame, lojas: pd.DataFrame, produtos: pd.DataFrame, resumo: pd.DataFrame, config_grade: pd.DataFrame = None, controle_extracao: pd.DataFrame = None) -> bytes:
    output = io.BytesIO()
    config_grade = preparar_config_grade(config_grade)
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        grade.to_excel(writer, sheet_name="NOTA_FINAL", index=False)
        base.to_excel(writer, sheet_name="BASE_LIMPA", index=False)
        validacoes.to_excel(writer, sheet_name="VALIDACOES", index=False)
        resumo.to_excel(writer, sheet_name="RESUMO_PEDIDOS", index=False)
        if controle_extracao is None:
            controle_extracao = pd.DataFrame()
        controle_extracao.to_excel(writer, sheet_name="CONTROLE_EXTRACAO", index=False)
        lojas.to_excel(writer, sheet_name="DE_PARA_LOJAS", index=False)
        produtos.to_excel(writer, sheet_name="DE_PARA_PRODUTOS", index=False)
        config_grade.to_excel(writer, sheet_name="CONFIG_GRADE", index=False)
        pd.DataFrame([
            [APP_NAME],
            ["Fluxo: subir PDF TOTVS > conferir validações > baixar NOTA_FINAL."],
            ["Atualize DE_PARA_LOJAS e DE_PARA_PRODUTOS sempre que aparecer loja/produto novo."],
            ["Use 'Usar preço referência' = SIM na loja que deve mandar o preço oficial da NOTA_FINAL."],
            ["Use CONFIG_GRADE para ocultar lojas ou alterar a posição das colunas."],
            ["Validações críticas devem ser corrigidas antes de enviar a grade."],
        ]).to_excel(writer, sheet_name="LEIA-ME", index=False, header=False)

        workbook = writer.book
        fmt_title = workbook.add_format({"bold": True, "font_size": 14, "font_color": "white", "bg_color": "#1F4E78", "align": "center", "valign": "vcenter"})
        fmt_header = workbook.add_format({"bold": True, "font_color": "white", "bg_color": "#1F4E78", "border": 1, "align": "center", "valign": "vcenter"})
        fmt_int = workbook.add_format({"num_format": "#,##0.00", "border": 1})
        fmt_money = workbook.add_format({"num_format": "R$ #,##0.00", "border": 1})
        fmt_text = workbook.add_format({"border": 1})
        fmt_crit = workbook.add_format({"bg_color": "#F4CCCC"})
        fmt_warn = workbook.add_format({"bg_color": "#FFF2CC"})
        fmt_ok = workbook.add_format({"bg_color": "#D9EAD3"})

        for sheet_name, df in [
            ("NOTA_FINAL", grade), ("BASE_LIMPA", base), ("VALIDACOES", validacoes),
            ("RESUMO_PEDIDOS", resumo), ("CONTROLE_EXTRACAO", controle_extracao),
            ("DE_PARA_LOJAS", lojas), ("DE_PARA_PRODUTOS", produtos),
            ("CONFIG_GRADE", config_grade)
        ]:
            ws = writer.sheets[sheet_name]
            ws.freeze_panes(1, 1)
            ws.autofilter(0, 0, max(len(df), 1), max(len(df.columns) - 1, 0))
            for col_idx, col_name in enumerate(df.columns):
                ws.write(0, col_idx, col_name, fmt_header)
                width = min(max(len(str(col_name)) + 2, 10), 35)
                if len(df) > 0:
                    try:
                        width = min(max(df[col_name].astype(str).map(len).max() + 2, width), 45)
                    except Exception:
                        pass
                money_cols = {"Preço unitário", "Valor do item", "Valor calculado", "Diferença item", "Total do pedido PDF", "Total_PDF", "Total_Itens", "Diferença", "PREÇO"}
                if col_name in money_cols:
                    ws.set_column(col_idx, col_idx, max(width, 13), fmt_money)
                else:
                    ws.set_column(col_idx, col_idx, width, fmt_text)

            if sheet_name == "NOTA_FINAL" and not df.empty:
                ws.set_column(0, 0, 34, fmt_text)
                for col_idx, col_name in enumerate(df.columns):
                    if col_name == "Produto":
                        ws.set_column(col_idx, col_idx, 34, fmt_text)
                    elif col_name == "PREÇO":
                        ws.set_column(col_idx, col_idx, 12, fmt_money)
                    else:
                        ws.set_column(col_idx, col_idx, 10, fmt_int)
                # Destaque para total e preço e aplicação de fórmula SOMA
                total_col = df.columns.get_loc("TOTAL") if "TOTAL" in df.columns else None
                preco_col = df.columns.get_loc("PREÇO") if "PREÇO" in df.columns else None
                for c in [total_col, preco_col]:
                    if c is not None:
                        ws.set_column(c, c, 13, fmt_money if df.columns[c] == "PREÇO" else fmt_int)
                
                # Substitui a coluna TOTAL da NOTA_FINAL por fórmulas de soma
                if total_col is not None and not df.empty:
                    last_store_col_idx = total_col - 1
                    start_letter = xl_col_to_name(1)  # B
                    end_letter = xl_col_to_name(last_store_col_idx)
                    for i in range(len(df)):
                        r_excel = i + 2  # dados começam na linha 2 física
                        formula = f"=SUM({start_letter}{r_excel}:{end_letter}{r_excel})"
                        ws.write_formula(i + 1, total_col, formula, fmt_int)

            if sheet_name == "VALIDACOES" and not df.empty:
                last_row = len(df)
                last_col = max(len(df.columns) - 1, 0)
                ws.conditional_format(1, 0, last_row, last_col, {"type": "text", "criteria": "containing", "value": "CRÍTICO", "format": fmt_crit})
                ws.conditional_format(1, 0, last_row, last_col, {"type": "text", "criteria": "containing", "value": "ATENÇÃO", "format": fmt_warn})
                ws.conditional_format(1, 0, last_row, last_col, {"type": "text", "criteria": "containing", "value": "OK", "format": fmt_ok})

        ws = writer.sheets["LEIA-ME"]
        ws.set_column(0, 0, 110)
        ws.merge_range("A1:D1", APP_NAME, fmt_title)
        ws.write("A3", "Como usar:")
        ws.write("A4", "1. Abra o sistema no Streamlit.")
        ws.write("A5", "2. Suba o PDF de pedidos TOTVS.")
        ws.write("A6", "3. Confira a aba VALIDACOES antes de usar a grade.")
        ws.write("A7", "4. Se aparecer loja/produto sem cadastro, atualize os DE/PARA e rode novamente.")
        
        # Cria as abas de cada loja
        adicionar_abas_lojas(writer, base, workbook)
    return output.getvalue()


def adicionar_abas_lojas(writer, base: pd.DataFrame, workbook) -> None:
    """Cria abas adicionais no arquivo Excel para cada loja que contiver pedidos."""
    if base is None or base.empty:
        return
    
    # Formatos específicos
    fmt_header = workbook.add_format({
        "bold": True,
        "font_color": "white",
        "bg_color": "#1F4E78",
        "border": 1,
        "align": "center",
        "valign": "vcenter"
    })
    fmt_text = workbook.add_format({"border": 1, "valign": "vcenter"})
    fmt_qty = workbook.add_format({"num_format": "#,##0.00", "border": 1, "align": "center", "valign": "vcenter"})
    fmt_money = workbook.add_format({"num_format": "R$ #,##0.00", "border": 1, "bold": True, "bg_color": "#FFF2CC", "align": "center"})
    fmt_total_label = workbook.add_format({"bold": True, "border": 1, "valign": "vcenter"})
    fmt_currency_symbol = workbook.add_format({"border": 1, "align": "center", "valign": "vcenter"})

    # Filtra apenas linhas com quantidade maior que zero
    base_filtrada = base[base["Quantidade"].fillna(0).gt(0)].copy()
    if base_filtrada.empty:
        return
    
    # Lojas presentes na base que têm quantidades
    lojas_presentes = base_filtrada["Coluna da grade"].dropna().unique()
    lojas_presentes = sorted([str(l).strip() for l in lojas_presentes if str(l).strip()])
    
    for loja in lojas_presentes:
        df_loja = base_filtrada[base_filtrada["Coluna da grade"] == loja]
        if df_loja.empty:
            continue
        
        # Agrupa por produto para somar quantidades e pega o primeiro preço unitário
        df_grouped = df_loja.groupby("Produto padronizado", as_index=False).agg({
            "Quantidade": "sum",
            "Preço unitário": "first"
        })
        df_grouped = df_grouped.sort_values("Produto padronizado").reset_index(drop=True)
        
        if df_grouped.empty:
            continue
        
        # Limpa o nome da aba para ser compatível com as restrições do Excel (máx 31 caracteres)
        import re
        sheet_name = re.sub(r"[\\*?:/\[\]]", "", loja)[:31].strip()
        if not sheet_name:
            continue
        
        ws = workbook.add_worksheet(sheet_name)
        ws.freeze_panes(1, 0)
        
        # Cabeçalhos
        ws.write(0, 0, "Produto", fmt_header)
        ws.write(0, 1, loja, fmt_header)
        ws.write(0, 2, "PREÇO", fmt_header)
        ws.set_row(0, 24)
        
        # Escreve os dados
        for idx, row in df_grouped.iterrows():
            r_excel = idx + 1
            ws.write(r_excel, 0, row["Produto padronizado"], fmt_text)
            ws.write(r_excel, 1, row["Quantidade"], fmt_qty)
            ws.write(r_excel, 2, row["Preço unitário"], fmt_money)
            ws.set_row(r_excel, 20)
            
        # Linha de total
        total_row = len(df_grouped) + 1
        ws.write(total_row, 0, "TOTAL PEDIDO", fmt_total_label)
        ws.write(total_row, 1, "R$", fmt_currency_symbol)
        
        # Fórmula SUMPRODUCT
        formula = f"=SUMPRODUCT(B2:B{total_row}, C2:C{total_row})"
        ws.write_formula(total_row, 2, formula, fmt_money)
        ws.set_row(total_row, 22)
        
        ws.set_column(0, 0, 34, fmt_text)
        ws.set_column(1, 1, 12, fmt_qty)
        ws.set_column(2, 2, 14, fmt_money)

