import pandas as pd
import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode

# Aumenta o limite de renderização do Styler
pd.set_option("styler.render.max_elements", 400000)

# Configuração da página
st.set_page_config(page_title="MRP - Estoque e Fluxo", layout="wide")
st.markdown("### 📦 Visão Consolidada de Estoques\Transações")

# Upload do arquivo Excel
uploaded_file = st.file_uploader("📁 Faça o upload da Base MRP (Excel)", type=["xlsx"])

if uploaded_file is not None:
    try:
        aba = 'report1'
        df = pd.read_excel(uploaded_file, sheet_name=aba)

        # === Estoque Atual ===
        df_estoque = df[df['Operação'] == '0 - Estoque Atual']
        estoque_por_item = df_estoque.groupby(['Item', 'Un'], as_index=False)['Estoque'].sum()
        estoque_por_item.rename(columns={'Estoque': 'Estoque Atual'}, inplace=True)

        # === Saídas ===
        df_saida = df[df['Operação'].str.contains("Saída", case=False, na=False)]
        df_saida = df_saida[df_saida['Data Corte Malha'].notna()]
        df_saida['Data Corte Malha'] = pd.to_datetime(df_saida['Data Corte Malha'])

        saida = df_saida.groupby(['Item', 'Data Corte Malha'])['Estoque'].sum().reset_index()
        saida['Estoque'] *= -1  # Saídas negativas
        saida_pivot = saida.pivot_table(index='Item', columns='Data Corte Malha', values='Estoque', fill_value=0)

        # === Entradas ===
        df_entrada = df[df['Operação'].str.contains("Entrada", case=False, na=False)]
        df_entrada = df_entrada[df_entrada['Data Entrega Malha'].notna()]
        df_entrada['Data Entrega Malha'] = pd.to_datetime(df_entrada['Data Entrega Malha'])

        entrada = df_entrada.groupby(['Item', 'Data Entrega Malha'])['Estoque'].sum().reset_index()
        entrada_pivot = entrada.pivot_table(index='Item', columns='Data Entrega Malha', values='Estoque', fill_value=0)

        # === Todas as datas ===
        todas_datas = sorted(set(entrada_pivot.columns).union(set(saida_pivot.columns)))

        # === Filtro ===
        itens_disponiveis = sorted(estoque_por_item['Item'].unique())
        item_opcao = [""] + itens_disponiveis  # primeira opção em branco
        item_selecionado = st.selectbox("🔍 Selecione um Item para visualizar (ou deixe em branco para ver todos)", item_opcao)

        if item_selecionado:
            itens_para_mostrar = [item_selecionado]
        else:
            total_paginas = (len(itens_disponiveis) - 1) // 500 + 1
            pagina = st.selectbox("🔢 Página de Itens", [f"Página {i+1}" for i in range(total_paginas)])
            idx = int(pagina.split()[1]) - 1
            itens_para_mostrar = itens_disponiveis[idx*500:(idx+1)*500]

        consolidado = estoque_por_item.set_index('Item')
        consolidado = consolidado.loc[consolidado.index.intersection(itens_para_mostrar)]
        consolidado['__saldo_atual__'] = consolidado['Estoque Atual']

        for data in todas_datas:
            col_entrada = f"{data.strftime('%d/%m/%Y')} Entrada"
            col_saida = f"{data.strftime('%d/%m/%Y')} (-) Saída"
            col_saldo = f"{data.strftime('%d/%m/%Y')} Saldo"

            entrada_val = entrada_pivot.get(data, pd.Series(0, index=consolidado.index))
            saida_val = saida_pivot.get(data, pd.Series(0, index=consolidado.index))

            consolidado[col_entrada] = entrada_val
            consolidado[col_saida] = saida_val

            consolidado[col_saldo] = consolidado['__saldo_atual__'] + entrada_val + saida_val
            consolidado['__saldo_atual__'] = consolidado[col_saldo]

        consolidado.drop(columns=['__saldo_atual__'], inplace=True)
        consolidado = consolidado.fillna(0).reset_index()

        # === Formatação ===
        def formatar_brasileiro(val):
            try:
                return f"{val:,.2f}".replace(",", "@").replace(".", ",").replace("@", ".")
            except:
                return val

        for col in consolidado.select_dtypes(include='number').columns:
            consolidado[col] = consolidado[col].apply(formatar_brasileiro)

        # === AGGRID ===
        gb = GridOptionsBuilder.from_dataframe(consolidado)
        gb.configure_default_column(resizable=True, filter=True, sortable=True)
        gb.configure_column("Item", pinned="left")

        # Estilo condicional para colunas de saldo: fundo cinza, vermelho se negativo
        style_saldo = JsCode('''
        function(params) {
            if (params.value != null) {
                let numero = parseFloat(params.value.replace(".", "").replace(",", "."));
                if (!isNaN(numero)) {
                    return {
                        'backgroundColor': '#f0f0f0',
                        'color': numero < 0 ? 'red' : 'black'
                    }
                }
            }
            return {
                'backgroundColor': '#f0f0f0'
            }
        }
        ''')

        for col in consolidado.columns:
            if "Saldo" in col:
                gb.configure_column(col, cellStyle=style_saldo)

        grid_options = gb.build()

        AgGrid(
            consolidado,
            gridOptions=grid_options,
            height=600,
            fit_columns_on_grid_load=False,
            enable_enterprise_modules=False,
            allow_unsafe_jscode=True
        )

        # Exportar
        csv = consolidado.to_csv(index=False).encode('utf-8')
        st.download_button("📅 Baixar CSV", data=csv, file_name='Estoque_Fluxo.csv', mime='text/csv')

    except Exception as e:
        st.error(f"Erro ao carregar os dados: {e}")
else:
    st.info("📝 Envie a planilha para visualizar os dados.")

