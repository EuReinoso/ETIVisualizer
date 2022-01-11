import pyodbc, csv, os
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib as mpl

DRIVER = "{Microsoft Access Driver (*.mdb, *.accdb)}"
SELECT_ETIQUETA = 'SELECT * FROM ETIQUETA;'
SELECT_LINHA = 'SELECT * FROM LINHA;'
ETIQUETA_COL_NAMES = ["ETI_CODIGO","ETI_NOME","ETI_PORTA","ETI_IMPRESSORA","ETI_DENS","ETI_VELOCIDADE","ETI_COLUNA","ETI_ALTURA","ETI_LARGURA","ETI_MARG_ESQ","ETI_MARG_CEN","ETI_ORIENTACAO","ETI_PROMOCAO","ETI_REDE","ETI_MASCARA","ETI_TEMPO","ETI_ENTREALTURA","ETI_COLUNAINVERTIDA","ETI_FORMATAVALOR","ETI_CARREGALOGO","ETI_MARG_ALT","ETI_ALTERARPRECO","ETI_NOMEIMPRESSORA","ETI_ARREDONDA","ETI_TEMPO_WINDOWS","ETI_ARQUIVO_TEMPLATE","ETI_TIPO_ETIQUETA","ETI_INATIVA","ETI_MATRICULA","ETI_ATUALIZACAO","ETI_USAR_CONFIG_IMPRESSORA"]
LINHA_COL_NAMES = ["LIN_ETIQUETA","LIN_TIPO","LIN_TEXTO","LIN_COLUNA","LIN_ALTURA","LIN_MARGEN","LIN_FONTE","LIN_ORIENTACAO","LIN_INICIO","LIN_FINAL","LIN_MASCARA","LIN_TABELA_PRECO"]
LIN_COLUNA_1 = {
    '0' : '',
    '1' : 'COD-PRODUTO',
    '2' : 'DESCRIÇÃO',
    '3' : 'GRUPO',
    '4' : 'COLEÇÃO',
    '5' : 'LINHA',
    '6' : 'MATERIAL',
    '7' : 'COR',
    '8' : 'TAM',
    '9' : 'FORNECEDOR',
    '10' : 'REF-FORNECEDOR',
    '11' : 'COD-PEDIDO',
    '13' : 'MODELO',
    '14' : 'REFERENCIA',
    '15' : 'COD-LOJA',
    '18' : 'LOJA',
    '19' : 'COD-FORNECEDOR',
    '22' : 'PREÇO-PADRAO',
    '23' : 'METRAGEM',
    '25' : 'COR-SEQ',
    '26' : 'ALTURA', 
    '28' : 'PROFUNDIDADE',
    '27' : 'LARGURA',
    '29' : 'ENDEREÇAMENTO',
    '30' : 'SALDO-ATUAL',
    '31' : 'COD-SECUNDÁRIO',
    '32' : 'DATA-DE-CADASTRO',
    '40' : '999,99',
    '41' : 'EAN',
    '42' : 'ULTIMA-COMPRA',
    '43' : 'VALIDADE',
    '44' : '999,99/100g',
    '201' : 'ESTADO-DO-PRODUTO',
}

LIN_FONTE_1 = {
    '0' : 6,
    '1' : 8,
    '2' : 12,
    '3' : 14,
    '4' : 18,
    '5' : 24,
}

#sistem ainda não comporta legenda
LIN_FONTE_2 = {
    '0' : (2, 0.2),
    '1' : (2, 0.2), 
    '2' : (2, 0.2),
    '3' : (2, 0.2),
    '4' : (3, 0.3),
    '5' : (3, 0.3),
    '6' : (3, 0.3),
    '7' : (3, 0.3),
    '8' : (4, 0.4),
    '9' : (4, 0.4),
    '10' : (4, 0.4),
    '11' : (4, 0.4),
    '12' : (3, 0.4),
    '13' : (3, 0.4),
    '14' : (2, 0.2),
    '15' : (2, 0.2),
}


eti_nome = None
last_eti = 0


#INICIO
os.system('cls')
print('IMPLANTAÇÃO - VISUALIZADOR DE ETIQUETA\n')

while True: 
    
    file_path = input('\nDigite o diretório do arquivo ETI.mdb: ')

    try:
        con = pyodbc.connect(r'Driver={};DBQ={};'.format(DRIVER, file_path))
        print('Database Conectado - OK!')
        break
    except pyodbc.Error as e:
        print("Erro ao conectar ao Database: ", e)
        print()
        print("- VERIFIQUE SE O DIRETORIO ESTÁ CORRETO")
        print("- VERIFIQUE SE OS DRIVERS DO MS ACCESS ESTÃO INSTALADOS")

while True:
    try:
        cur = con.cursor()
        table_etiqueta = cur.execute(SELECT_ETIQUETA).fetchall()
        table_linha = cur.execute(SELECT_LINHA).fetchall()

        with open('ETIQUETA.csv', 'w') as fou:
            csv_writer = csv.writer(fou)
            csv_writer.writerows(table_etiqueta)

        with open('LINHA.csv', 'w') as fou:
            csv_writer = csv.writer(fou)
            csv_writer.writerows(table_linha)

        print('Tabelas CSV criadas - OK!')

    except:
        print('Erro ao CRIAR Tabelas CSV.')
        exit()


    try:
        df_etiqueta = pd.read_csv('ETIQUETA.csv', names=ETIQUETA_COL_NAMES)
        df_linha = pd.read_csv('LINHA.csv', names=LINHA_COL_NAMES)
        print('Leitura do CSV - OK!')

    except:
        print('Erro ao LER tabelas CSV.')


    while eti_nome == None:

        eti_nome = input("\nDigite o NOME da etiqueta: ")

        if len(df_etiqueta.loc[df_etiqueta['ETI_NOME'] == eti_nome].values) > 0:
            print("Etiqueta Encontrada - OK!")
            break

        else:
            print('\nERRO - Etiqueta NÃO encontrada!')
            print('- VEFIQUE SE O NOME DIGITADO CORRESPONDE AO DO SISTEMA (RESPEITE ESPAÇOS E ACENTOS)')
            print('- VERIFIQUE SE O CAMPO LIN_CODIGO E LIN_NOME DO ARQUIVO MDB ESTÃO TROCADOS')

    #LEITURA DE PARAMETROS DA ETI
    eti_cod      = df_etiqueta.loc[df_etiqueta['ETI_NOME'] == eti_nome]['ETI_CODIGO'].values[0]
    eti_altura   = df_etiqueta.loc[df_etiqueta['ETI_NOME'] == eti_nome]['ETI_ALTURA'].values[0]
    eti_largura  = df_etiqueta.loc[df_etiqueta['ETI_NOME'] == eti_nome]['ETI_LARGURA'].values[0]
    eti_marg_esq = df_etiqueta.loc[df_etiqueta['ETI_NOME'] == eti_nome]['ETI_MARG_ESQ'].values[0]
    eti_marg_cen = df_etiqueta.loc[df_etiqueta['ETI_NOME'] == eti_nome]['ETI_MARG_CEN'].values[0]

    orien = 1

    eti_linhas = df_linha.loc[df_linha['LIN_ETIQUETA'] == eti_cod]

    #LOOP PRINCIPAL
    while True:
        print('\nMENU')
        print('1 - Ver Etiqueta')
        print('2 - Atualizar Etiqueta')
        print('3 - Mudar Orientação')
        print('4 - Sair')
        i = input("Escolha a opção: ")

        if i == '1':
            fig, ax = plt.subplots(figsize=(eti_largura, eti_altura), dpi=150)

            ax.set_facecolor('#cccccc')


            plt.xlim(0 - eti_marg_esq, eti_largura + eti_marg_cen)

            if orien == 1:
                plt.ylim(0, eti_altura)
            elif orien == 0:
                plt.ylim(eti_altura, 0)


            rect = mpl.patches.Rectangle((0, 0), eti_largura, eti_altura, color='white')

            df = eti_linhas

            #COLUNAS
            rows = len(df.index)
            for i in range(rows):
                if df.iloc[i]['LIN_TIPO'] == 1:
                    col = df.iloc[i]['LIN_COLUNA']
                    size = df.iloc[i]['LIN_FONTE']
                    x = df.iloc[i]['LIN_MARGEN']
                    y = df.iloc[i]['LIN_ALTURA']
                    font_size = LIN_FONTE_1[str(size)]
                    marg = 0
                    if pd.notnull(df.iloc[i]['LIN_TEXTO']):
                        text = df.iloc[i]['LIN_TEXTO'].replace(' ', '   ')
                        marg = len(text) / 8
                        ax.text(x, y, '{}'.format(text),size=font_size)

                    text = LIN_COLUNA_1[str(col)]
                    ax.text(x + marg, y, '{}'.format(text),size=font_size)

                
                if df.iloc[i]['LIN_TIPO'] == 2:
                    x = df.iloc[i]['LIN_MARGEN']
                    y = df.iloc[i]['LIN_ALTURA']
                    width = x + LIN_FONTE_2[str(df.iloc[i]['LIN_FONTE'])][0]
                    height = y + LIN_FONTE_2[str(df.iloc[i]['LIN_FONTE'])][1]
                    
                    barcode_img = plt.imread('barcode.png')

                    plt.imshow(barcode_img, extent=[x, width, y, height], aspect='auto', zorder=2)

            ax.add_patch(rect)
            plt.savefig('IMAGENS/eti_{}_{}.jpg'.format(eti_nome, last_eti))
            last_eti += 1
            
            os.system('cls')
            print("\nIMAGEM BAIXADA!")
        elif i == '2':
            os.system('cls')
            print('\nATUALIZANDO ETIQUETA...')
            break
        
        elif i == '3':
            if orien == 1:
                orien = 0
            else:
                orien = 1

            os.system('cls')
            print('\nORIENTAÇÃO MUDADA!')

        elif i == '4':
            exit()
    

