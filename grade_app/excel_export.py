"""Geracao dos arquivos Excel: NOTA FINAL enxuta e relatorio completo."""
import io
import pandas as pd
from xlsxwriter.utility import xl_col_to_name
from .utils import eh_sim
from .config import APP_NAME, preparar_config_grade


def _formatar_e_ajustar_nota_final(ws, df: pd.DataFrame, workbook, total_pedido_dia: float = 0.0, cliente_nome: str = ""):
    """Aplica estilos e insere fórmulas dinâmicas na aba NOTA_FINAL."""
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
    
    ws.write(1, max(last_col - 1, 0), "TOTAL DO PEDIDO DO DIA")
    
    total_col_idx = df.columns.get_loc("TOTAL") if "TOTAL" in df.columns else None
    preco_col_idx = df.columns.get_loc("PREÇO") if "PREÇO" in df.columns else None
    
    if len(df) > 0 and total_col_idx is not None and preco_col_idx is not None:
        letra_total = xl_col_to_name(total_col_idx)
        letra_preco = xl_col_to_name(preco_col_idx)
        formula_total = f"=SUMPRODUCT({letra_total}4:{letra_total}{len(df) + 3}, {letra_preco}4:{letra_preco}{len(df) + 3})"
        ws.write_formula(1, last_col, formula_total, fmt_money)
    else:
        ws.write(1, last_col, total_pedido_dia or 0.0, fmt_money)
        
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

    # Injetar fórmulas de SUM na coluna TOTAL para que sejam recalculados dinamicamente
    if len(df) > 0 and total_col_idx is not None:
        letra_ultima_loja = xl_col_to_name(total_col_idx - 1)
        for row_idx in range(3, len(df) + 3):
            linha_excel = row_idx + 1
            formula = f"=SUM(B{linha_excel}:{letra_ultima_loja}{linha_excel})"
            ws.write_formula(row_idx, total_col_idx, formula, fmt_total)

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
        _formatar_e_ajustar_nota_final(ws, df, workbook, total_pedido_dia, cliente_nome)
        
        # Criação de abas por loja
        adicionar_abas_lojas(writer, base, workbook)

    return output.getvalue()


def gerar_excel(base: pd.DataFrame, grade: pd.DataFrame, validacoes: pd.DataFrame, lojas: pd.DataFrame, produtos: pd.DataFrame, resumo: pd.DataFrame, config_grade: pd.DataFrame = None, controle_extracao: pd.DataFrame = None, total_pedido_dia: float = 0.0, cliente_nome: str = "") -> bytes:
    """Gera o arquivo Excel completo com todas as abas de conferencia e de/para."""
    output = io.BytesIO()
    config_grade = preparar_config_grade(config_grade)
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        grade.to_excel(writer, sheet_name="NOTA_FINAL", index=False, startrow=2)
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
        fmt_money = workbook.add_format({"num_format": "R$ #,##0.00", "border": 1})
        fmt_text = workbook.add_format({"border": 1})
        fmt_crit = workbook.add_format({"bg_color": "#F4CCCC"})
        fmt_warn = workbook.add_format({"bg_color": "#FFF2CC"})
        fmt_ok = workbook.add_format({"bg_color": "#D9EAD3"})

        # Formatar a aba NOTA_FINAL individualmente usando a função auxiliar
        ws_nota = writer.sheets["NOTA_FINAL"]
        _formatar_e_ajustar_nota_final(ws_nota, grade, workbook, total_pedido_dia, cliente_nome)

        # Outras abas (removido NOTA_FINAL dessa lista comum)
        for sheet_name, df in [
            ("BASE_LIMPA", base), ("VALIDACOES", validacoes),
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
                money_cols = {"Preço unitário", "Valor do item", "Valor calculated", "Diferença item", "Total do pedido PDF", "Total_PDF", "Total_Itens", "Diferença", "PREÇO"}
                if col_name in money_cols:
                    ws.set_column(col_idx, col_idx, max(width, 13), fmt_money)
                else:
                    ws.set_column(col_idx, col_idx, width, fmt_text)

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

        # Criação de abas por loja
        adicionar_abas_lojas(writer, base, workbook)

    return output.getvalue()


def adicionar_abas_lojas(writer, base: pd.DataFrame, workbook):
    """Cria uma nova aba no Excel para cada loja que contiver pedidos com quantidade acumulada > 0."""
    if base is None or base.empty:
        return

    # Verificar se as colunas necessárias existem
    necessarias = {"Coluna da grade", "Produto padronizado", "Quantidade", "Preço unitário"}
    if not necessarias.issubset(base.columns):
        return

    # 1. Calcular o preço unitário unificado usando a nova lógica de fallback
    # Para evitar dependências circulares de imports, replicamos a lógica direta
    base_comum = base[~base["Usar preço referência"].apply(eh_sim) & base["Preço unitário"].gt(0)]
    precos_comuns = base_comum.groupby("Produto padronizado")["Preço unitário"].first()

    if "Usar preço referência" in base.columns:
        base_preco_ref = base[base["Usar preço referência"].apply(eh_sim) & base["Preço unitário"].gt(0)].copy()
    else:
        base_preco_ref = pd.DataFrame(columns=base.columns)

    if not base_preco_ref.empty:
        precos_ref = base_preco_ref.groupby("Produto padronizado")["Preço unitário"].first()
    else:
        precos_ref = pd.Series(dtype=float)

    precos_final = precos_comuns.combine_first(precos_ref)

    # 2. Agrupar quantidades por loja (Coluna da grade) e produto
    df_valid = base[base["Coluna da grade"].fillna("").str.strip().ne("")].copy()
    if df_valid.empty:
        return

    df_agrupado = df_valid.groupby(["Coluna da grade", "Produto padronizado"])["Quantidade"].sum().reset_index()

    # 3. Identificar as lojas que têm quantidade acumulada > 0
    soma_lojas = df_valid.groupby("Coluna da grade")["Quantidade"].sum()
    lojas_com_pedidos = soma_lojas[soma_lojas > 0].index.tolist()

    # 4. Estilos de formatação para as abas de loja
    fmt_store_header = workbook.add_format({
        "bold": True,
        "font_color": "white",
        "bg_color": "#1F4E78",
        "border": 1,
        "align": "center",
        "valign": "vcenter"
    })
    fmt_store_text = workbook.add_format({
        "border": 1,
        "valign": "vcenter"
    })
    fmt_store_qty = workbook.add_format({
        "num_format": "#,##0.00",
        "border": 1,
        "align": "center",
        "valign": "vcenter"
    })
    fmt_store_money = workbook.add_format({
        "num_format": "R$ #,##0.00",
        "border": 1,
        "valign": "vcenter"
    })
    fmt_tot_label = workbook.add_format({
        "bold": True,
        "border": 1,
        "align": "left",
        "valign": "vcenter"
    })
    fmt_tot_rs = workbook.add_format({
        "border": 1,
        "align": "center",
        "valign": "vcenter"
    })
    fmt_tot_formula = workbook.add_format({
        "num_format": "R$ #,##0.00",
        "border": 1,
        "bold": True,
        "bg_color": "#FFF2CC",
        "align": "center",
        "valign": "vcenter"
    })

    # 5. Iterar e criar as abas
    for loja in sorted(lojas_com_pedidos):
        sheet_name = str(loja)[:31] # O nome da aba no Excel tem limite de 31 caracteres

        # Filtra os dados da loja
        df_loja = df_agrupado[(df_agrupado["Coluna da grade"] == loja) & (df_agrupado["Quantidade"] > 0)].copy()
        df_loja["PREÇO"] = df_loja["Produto padronizado"].map(precos_final).fillna(0)
        df_loja = df_loja.sort_values("Produto padronizado").reset_index(drop=True)

        ws = workbook.add_worksheet(sheet_name)
        writer.sheets[sheet_name] = ws

        # Cabeçalhos: Produto (coluna A), Nome_Abreviado_Loja (coluna B) e PREÇO (coluna C)
        ws.write(0, 0, "Produto", fmt_store_header)
        ws.write(0, 1, str(loja), fmt_store_header)
        ws.write(0, 2, "PREÇO", fmt_store_header)

        # Ajustar larguras de colunas
        ws.set_column(0, 0, 34, fmt_store_text)
        ws.set_column(1, 1, 15, fmt_store_qty)
        ws.set_column(2, 2, 13, fmt_store_money)
        ws.set_row(0, 24)

        # Dados
        for i, row in df_loja.iterrows():
            linha_xlsx = i + 1
            ws.write(linha_xlsx, 0, row["Produto padronizado"], fmt_store_text)
            ws.write(linha_xlsx, 1, float(row["Quantidade"]), fmt_store_qty)
            ws.write(linha_xlsx, 2, float(row["PREÇO"]), fmt_store_money)
            ws.set_row(linha_xlsx, 20)

        # Linha final de totalizador da loja
        total_row_idx = len(df_loja) + 1
        ws.write(total_row_idx, 0, "TOTAL PEDIDO", fmt_tot_label)
        ws.write(total_row_idx, 1, "R$", fmt_tot_rs)

        formula = f"=SUMPRODUCT(B2:B{total_row_idx}, C2:C{total_row_idx})"
        ws.write_formula(total_row_idx, 2, formula, fmt_tot_formula)
        ws.set_row(total_row_idx, 22)

        ws.freeze_panes(1, 1)

