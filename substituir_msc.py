import streamlit as st
import pandas as pd
import io
import openpyxl

st.title("Atualizador de Arquivo MSC")

# Upload dos arquivos
uploaded_csv = st.file_uploader("Faça upload do arquivo da MSC atual no formato .CSV", type=["csv"])
uploaded_csv_anterior = st.file_uploader("Faça o upload da MSC anterior no formato .CSV", type=["csv"])
uploaded_xlsx = st.file_uploader("Faça upload do arquivo de distribuição por fontes no formato .XLSX", type=["xlsx"])
troca_saldo_inicial = False
if uploaded_csv_anterior:
    troca_saldo_inicial = True

if uploaded_csv_anterior and uploaded_csv:
    if uploaded_csv_anterior.getvalue() == uploaded_csv.getvalue():
        st.warning(
            "A MSC atual e a MSC anterior parecem ser idênticas. "
            "Verifique se os arquivos estão corretos."
        )

# Só mostra o botão se os dois arquivos (MSC e distribuição) forem enviados, MSC anterior não é obrigatório
if uploaded_csv and uploaded_xlsx:
    # Lista todas as abas do XLSX
    xls = pd.ExcelFile(uploaded_xlsx)
    sheet_names = xls.sheet_names
    uploaded_xlsx.seek(0)  # reseta o ponteiro do arquivo
    
    st.success(f"Arquivos carregados com sucesso! ✅\n\nForam detectadas as abas: {', '.join(sheet_names)}")
    if troca_saldo_inicial:
        st.success(f"Os saldos iniciais serão substituídos conforme os saldos finais da MSC anterior.")

    fechar_periodo = st.checkbox("Deseja fechar o período (baixar fontes não tratadas do mês anterior)?")

    if st.button("Confirmar e processar"):
        
        lista_erros = []

        # Lê os CSV em memória e transforma em listas:
        msc_lista = uploaded_csv.read().decode("utf-8").splitlines()
        msc_anterior = []
        if uploaded_csv_anterior:
            msc_anterior = uploaded_csv_anterior.read().decode("utf-8").splitlines()
        msc_nova = msc_lista.copy()
        itens_processados = []

        # Marca tudo que já foi tratado para depois comparar com os saldos finais da MSC anterior
        tratados = set()

        for item in msc_lista:
            partes = item.split(";")

            if len(partes) < 16:
                continue

            conta = partes[0]
            PO = partes[1]
            fonte = partes[5]
            fr = partes[6]
            tipo = partes[14]

            if tipo == "ending_balance" and fr == "FR" and fonte:
                tratados.add((conta, PO, fonte))

        # =====================
        # Passo 1 - xlsx
        # =====================

        # Itera sobre cada aba (conta)
        for conta in sheet_names: # Início do tratamento das abas do xlsx
            df = pd.read_excel(uploaded_xlsx, sheet_name=conta, dtype=str)

            # Definir a natureza da conta e o indicador FP
            if conta.startswith('1'):
                natureza_saldo_conta = 'D'
                natureza_mov_baixa = 'C'
                indicador_FP = "1;FP"
            elif conta.startswith('2'):
                natureza_saldo_conta = 'C'
                natureza_mov_baixa = 'D'
                indicador_FP = "1;FP"
            elif conta.startswith('8'):
                natureza_saldo_conta = 'C'
                natureza_mov_baixa = 'D'
                indicador_FP = ";"

            
            # POs únicos nessa aba
            pos_unicos = df.iloc[:, 1].dropna().unique()  # coluna 1 (segunda) é o PO

            for PO in pos_unicos:
                # Classifica todos os itens da MSC atual que atendem a combinação de conta e PO
                
                item_saldo_inicial = []
                item_movimento_baixa = []
                item_movimento_normal = []
                item_saldo_final = []
                itens_novos = []
               
                invertido = False
                
                for item in msc_lista:
                    if item.startswith(f'{conta};{PO};PO;{indicador_FP};') and item.endswith(f'beginning_balance;{natureza_saldo_conta}'): # saldo inicial normal
                        item_saldo_inicial.append(item)
                    elif item.startswith(f'{conta};{PO};PO;{indicador_FP};') and item.endswith(f'beginning_balance;{natureza_mov_baixa}'): # saldo inicial invertido
                        item_saldo_inicial.append(item)
                    elif item.startswith(f'{conta};{PO};PO;{indicador_FP};') and item.endswith(f'period_change;{natureza_saldo_conta}'): # movimento normal
                        item_movimento_normal.append(item)
                    elif item.startswith(f'{conta};{PO};PO;{indicador_FP};') and item.endswith(f"period_change;{natureza_mov_baixa}"): # movimento baixa
                        item_movimento_baixa.append(item)
                    elif item.startswith(f'{conta};{PO};PO;{indicador_FP};') and item.endswith(f'ending_balance;{natureza_saldo_conta}'): # saldo final normal 
                        item_saldo_final.append(item)
                        invertido = False
                    elif item.startswith(f'{conta};{PO};PO;{indicador_FP};') and item.endswith(f'ending_balance;{natureza_mov_baixa}'): # saldo final invertido
                        item_saldo_final.append(item)
                        invertido = True # Flag para indicar que o saldo final está invertido na MSC
                                      
                if not item_saldo_final:
                    erro = f"Aba {conta}, PO {PO} não encontrou linhas correspondentes na MSC atual. As linhas não foram geradas."
                    st.warning(erro)
                    lista_erros.append(erro)
                    continue
                
                if len(item_saldo_final) > 1:
                    erro = (
                        f"Aba {conta}, PO {PO}: "
                        f"foram encontrados {len(item_saldo_final)} saldos finais na MSC atual. As linhas não foram geradas."
                    )
                    st.warning(erro)
                    lista_erros.append(erro)
                    continue

                valor_end = float(item_saldo_final[0].split(";")[13])

                if not msc_anterior: # baixa dos saldos finais e substituição dos valores de mov das linhas SEM detalhamento de fonte, pois nesse caso as linhas
                                     # serão simplesmente apagadas, já que não partiremos mais do mesmo valor inicial.

                    # Verifica o valor da linha de movimento de baixa 
                    if not item_movimento_baixa:
                        valor_mov_baixa = 0
                    else:
                        valor_mov_baixa = float(item_movimento_baixa[0].split(";")[13])

                    # Verifica o valor da linha de movimento normal
                    if not item_movimento_normal:
                        valor_mov_normal = 0
                    else:
                        valor_mov_normal = float(item_movimento_normal[0].split(";")[13])
                    
                    # Calcula os novos valores:
                    valor_mov_baixa_novo = valor_mov_baixa + valor_end # O que já estava baixando + o saldo final que restava
                    valor_mov_normal_novo = valor_mov_normal + valor_end # Se o saldo final estiver invertido, deve se somar ao mov normal

                    # Substitui o valor de movimento de baixa             
                    partes = item_saldo_final[0].split(";")
                    partes[14] = 'period_change'

                    if partes[15] == natureza_saldo_conta: # Verifica a naturezas do saldo final, se for devedor, o movimento de baixa deve ser credor e vice-versa
                        partes[13] = f"{valor_mov_baixa_novo:.2f}"
                        partes[15] = natureza_mov_baixa
                    else:
                        partes[13] = f"{valor_mov_normal_novo:.2f}"
                        partes[15] = natureza_saldo_conta

                    itens_novos.append(";".join(partes)) # Cria a linha de movimento com o valor novo
                
                # Filtra só as linhas desse PO
                linhas_po = df[df.iloc[:, 1] == PO].values.tolist()

                # Verifica se a soma dos valores bate com a MSC atual
                linhas_po_df = df[df.iloc[:, 1] == PO]
                soma_valores = abs(float(pd.to_numeric(linhas_po_df.iloc[:, 3], errors="coerce").sum()))
                if round(soma_valores, 2) != round(valor_end, 2):
                    erro = f"Aba {conta}: O valor total do PO {PO} \(R\$ {soma_valores:.2f}\) não bate com o saldo final na MSC (R$ {valor_end})!"
                    st.warning(erro)
                    lista_erros.append(erro)    

                valores_por_fonte = {}

                # Cálculo do movimento e saldo final das linhas COM abertura de fonte:
                for linha in linhas_po:
                    fonte = linha[2]
                    valor = abs(float(linha[3])) # considerar valores absolutos e ajustar apenas a natureza da conta

                    valores_por_fonte[fonte] = float(linha[3])

                    if valor == 0:
                        erro = f"Aba {conta}, PO {PO}, fonte {fonte}: valor zero no XLSX foi ignorado."
                        st.warning(erro)
                        lista_erros.append(erro)
                        continue

                    chave_fonte = (conta, PO, fonte) 
                    if chave_fonte in tratados:
                        erro = f"Aba {conta}, PO {PO}: fonte {fonte} aparece mais de uma vez no XLSX. As linhas não foram geradas." 
                        st.warning(erro) 
                        lista_erros.append(erro)
                        continue
                    
                    tratados.add((conta, PO, fonte))

                    partes = item_saldo_final[0].split(";")
                    partes[5] = fonte
                    partes[6] = 'FR'
                    partes[13] = f'{float(valor):.2f}'
                    partes[14] = 'period_change'
                    if float(linha[3]) > 0:
                        partes[15] = natureza_saldo_conta
                        invertido_mes_atual = False
                    else:
                        partes[15] = natureza_mov_baixa # ajustar a natureza quando o valor informado no xls for negativo
                        invertido_mes_atual = True
                    if not msc_anterior:                        # Cria as linhas de movimento para gerar saldo em cada fonte (pelo mesmo valor do saldo final), caso
                        itens_novos.append(";".join(partes))    # contrário, cria apenas a linha de saldo final
                    else:
                        item_saldo_final_mes_anterior = []
                        for item in msc_anterior:
                            if item.startswith(f'{conta};{PO};PO;{indicador_FP};{fonte};FR') and item.endswith(f'ending_balance;{natureza_saldo_conta}'): # saldo final normal 
                                item_saldo_final_mes_anterior.append(item)
                                invertido_mes_anterior = False
                            elif item.startswith(f'{conta};{PO};PO;{indicador_FP};{fonte};FR') and item.endswith(f'ending_balance;{natureza_mov_baixa}'): # saldo final invertido
                                item_saldo_final_mes_anterior.append(item)
                                invertido_mes_anterior = True # Flag para indicar que o saldo final estava invertido na MSC do mês anterior
                        if not item_saldo_final_mes_anterior:
                            itens_novos.append(";".join(partes))
                    partes[14] = 'ending_balance'
                    itens_novos.append(";".join(partes)) # Cria as linhas de saldo final para cada fonte

                    if not fonte or str(fonte).strip() == "":
                        erro = f"Aba {conta}, PO {PO}: existe linha no XLSX sem fonte definida."
                        st.warning(erro)
                        lista_erros.append(erro)
                        continue
                
                if msc_anterior:
                    
                    # Lista as fontes desse PO
                    fontes_unicas_no_po = linhas_po_df.iloc[:, 2].dropna().unique()

                    for fonte in fontes_unicas_no_po:

                        saldo_final_fonte = f'{valores_por_fonte[fonte]:.2f}' # vem do xlsx!
                        if abs(float(saldo_final_fonte)) == 0:
                            continue

                        # Localiza o saldo final da MSC anterior
                        item_saldo_final_mes_anterior = []

                        for item in msc_anterior:
                            if item.startswith(f'{conta};{PO};PO;{indicador_FP};{fonte};FR') and item.endswith(f'ending_balance;{natureza_saldo_conta}'): # saldo final normal 
                                item_saldo_final_mes_anterior.append(item)
                                invertido_mes_anterior = False
                            elif item.startswith(f'{conta};{PO};PO;{indicador_FP};{fonte};FR') and item.endswith(f'ending_balance;{natureza_mov_baixa}'): # saldo final invertido
                                item_saldo_final_mes_anterior.append(item)
                                invertido_mes_anterior = True # Flag para indicar que o saldo final estava invertido na MSC do mês anterior

                        # Transforma em saldo inicial da MSC atual
                        if item_saldo_final_mes_anterior:
                            partes = item_saldo_final_mes_anterior[0].split(";")
                            partes[14] = "beginning_balance"
                            itens_novos.append(";".join(partes)) # Cria a linhas de saldo inicial para cada fonte
                            if invertido_mes_anterior:
                                saldo_inicial_fonte = f'{-1*float(partes[13]):.2f}'
                            else:
                                saldo_inicial_fonte = f'{float(partes[13]):.2f}'
                        else:
                            saldo_inicial_fonte = 0

                        

                        saldo_mov_fonte = float(saldo_final_fonte) - float(saldo_inicial_fonte) # pode ser negativo
                        
                        if saldo_mov_fonte < 0:
                            natureza = natureza_mov_baixa
                            sem_movimento = False
                        elif saldo_mov_fonte > 0:
                            natureza = natureza_saldo_conta
                            sem_movimento = False
                        else:
                            sem_movimento = True # MSC não pode ter linhas zeradas!

                        if item_saldo_final_mes_anterior and not sem_movimento:
                            partes[13] = f"{abs(saldo_mov_fonte):.2f}"
                            partes[14] = 'period_change'
                            partes[15] = natureza
                            itens_novos.append(";".join(partes)) # Cria a linha de movimento para cada fonte    

                # Substitui no resultado final
                nova_lista = []

                if not msc_anterior:
                    for item in msc_nova:
                        # Saldo inicial → mantém sempre como estava
                        if item in item_saldo_inicial:
                            nova_lista.append(item)
                        # Saldo final → apaga e substitui pelos itens da lista itens_novos
                        elif item in item_saldo_final:
                            for item_novo in itens_novos:
                                nova_lista.append(item_novo)
                        # Movimento baixa
                        elif item in item_movimento_baixa: # mantém apenas se o saldo final estiver invertido na MSC, pois é o mov normal que deve mudar
                            if invertido:
                                nova_lista.append(item)
                        # Movimento normal
                        elif item in item_movimento_normal: # mantém o mov normal e modifica o de baixa, caso o saldo final NÃO esteja invertido na MSC
                            if not invertido:
                                nova_lista.append(item)
                        # TODAS AS OUTRAS LINHAS
                        else:
                            nova_lista.append(item)

                elif msc_anterior:
                    for item in msc_nova:
                        if item in item_saldo_inicial:
                            pass # Os itens de saldo inicial serão substituidos pelos gerados a partir do saldo final da MSC anterior
                        elif item in item_movimento_normal:
                            pass # Os itens de movimento serão calculados pela diferença entre saldo inicial e saldo final
                        elif item in item_movimento_baixa:
                            pass # Os itens de movimento serão calculados pela diferença entre saldo inicial e saldo final
                        elif item in item_saldo_final:
                            for item_novo in itens_novos:
                                nova_lista.append(item_novo) # Ignora o saldo final atual e gera todas as linhas com os itens novos calculados
                        else:
                            nova_lista.append(item) # Mantém todos os outros itens

                msc_nova = nova_lista

                chave_po = f"{conta}/{PO}"
                if chave_po in itens_processados:
                    erro = f"{chave_po} foi processado mais de uma vez."
                    st.warning(erro)
                    lista_erros.append(erro)
                else:
                    itens_processados.append(chave_po)

        # ============================
        # PASSO 2 - Varre MSC anterior
        # ============================

        if fechar_periodo and msc_anterior:

            baixas_passo2 = []

            for item in msc_anterior:
                partes = item.split(";")

                conta = partes[0]
                PO = partes[1]
                indicador_num = partes[3]   # "1"
                indicador_fp = partes[4]    # "FP"
                fonte = partes[5]
                fr = partes[6]
                tipo = partes[14]

                # 1) apenas saldo final
                if tipo != "ending_balance":
                    continue

                # 2) precisa ter detalhamento de fonte
                if not fonte or fr != "FR":
                    continue

                # 3) regras por tipo de conta
                # contas 1 e 2 → FP = 1
                if conta.startswith(("1", "2")):
                    if indicador_num != "1":
                        continue

                # contas 8 → não exige FP
                elif conta.startswith("8"):
                    pass

                # qualquer outra conta
                else:
                    continue

                # 👉 SE CHEGOU AQUI, A LINHA É ELEGÍVEL
                # (no Passo 3 vamos verificar se foi tratada e zerar)

                chave = (conta, PO, fonte)

                # se foi tratada pelo XLSX, ignora
                if chave in tratados:
                    continue

                # 👉 SE CHEGOU AQUI, É ELEGÍVEL E NÃO FOI TRATADA
                # (no Passo 4 vamos gerar saldo inicial + baixa)

                # valor e natureza do saldo final anterior
                valor = float(partes[13])
                natureza = partes[15]

                # natureza oposta para baixar
                natureza_baixar = "C" if natureza == "D" else "D"

                # --- saldo inicial ---
                partes_si = partes.copy()
                partes_si[14] = "beginning_balance"
                partes_si[13] = f"{valor:.2f}"
                baixas_passo2.append(";".join(partes_si))

                # --- movimento de baixa ---
                partes_mov = partes.copy()
                partes_mov[14] = "period_change"
                partes_mov[13] = f"{valor:.2f}"
                partes_mov[15] = natureza_baixar

                baixas_passo2.append(";".join(partes_mov))

            msc_nova.extend(baixas_passo2)
            
        # ========================
        # FINAL do passo 2
        # ========================
        
        cabecalho = msc_nova[:2]
        dados = msc_nova[2:]
        dados.sort()
        msc_nova = cabecalho + dados

        # Gera os arquivos em memória
        output = io.StringIO()
        output.write("\n".join(msc_nova))
        output.seek(0)

        erros = io.StringIO()
        erros.write("\n".join(lista_erros))
        erros.seek(0)

        st.success(f"Processamento concluído! Contas/POs processados: {', '.join(itens_processados)}")

        # Botão de download
        st.download_button(
            label="Baixar MSC atualizada",
            data=output.getvalue(),
            file_name="MSC_atualizada.csv",
            mime="text/csv"
        )

        if lista_erros:
            st.download_button(
                label="Baixar log de erros",
                data=erros.getvalue(),
                file_name="erros.txt",
                mime="text/csv"
            )
