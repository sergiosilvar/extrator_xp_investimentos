import locale
import pandas as pd
import re
from pathlib import Path
import logging
import numpy as np
import xlrd
import os



logger = logging.getLogger(__name__)

try:
    locale.setlocale( locale.LC_ALL, 'pt_BR' )
except Exception:
    locale.setlocale( locale.LC_ALL, '' )

_re_valor = r'\d+(?:\.\d{1,2})?'
_regex_valor = re.compile('(?P<tipo>LCA|LCI|LC|CDB)\s(\w+).+(?:\w+).+(\d{19})\s+('+_re_valor+').+(\d{4}-\d{2}-\d{2})\s00:00:00\s+(' + 
    _re_valor+')\s+('+_re_valor+')\s+('+_re_valor+')\s+('+_re_valor+')')
_regex_mesref = re.compile(r'DATA DE REFERÊNCIA:\s(\d{2}/\d{2}/\d{4})')

# BASE_DIR é uma variável de ambiente para a pasta onde estão os extratos.
BASE_DIR = ''
try:
    BASE_DIR = os.environ['EXTRATO']
except KeyError:
    print('Variavel de ambiente "EXTRATO" não existe! Defina com o local dos extratos.')


def _path(pasta):
    path = Path(pasta)
    if not path.exists():
         raise ValueError('Pasta {} nao existe!'.format(path.as_posix()))
    return path


def _CSV_DIR_POSICAO():
    return BASE_DIR +'/XP Posicao'

def _CSV_DIR_EXTRATO():
    return  BASE_DIR+'/XP'

def compila_posicao(pasta=None):
    '''
    Compila extratos de posicao dos investimento no formato XLS dos site para um objeto
    pandas.DataFrame.

    Parâmetros
    ----------
    pasta - Pasta com todos os arquivos CSV.
    '''

    if not pasta:
        pasta = _CSV_DIR_POSICAO()
    path = _path(pasta)

    # Não há problemas em converter para uma lista, não serão muitos arquivos.
    csv_files = list(path.glob('*.xls'))
    logger.debug('{} arquivos CSV encontrados em {}.'.format(len(csv_files),
        path.absolute().as_posix()))
    
    lista = []
    for i in csv_files:
        arquivo = i.name
        wb = xlrd.open_workbook(i.as_posix(), logfile=open(os.devnull, 'w'))
        xdf = pd.read_excel(wb, engine='xlrd', convert_float=False, dtype='str')
        #xdf = pd.read_excel(i.as_posix(), convert_float=False, engine='xlrd', dtype='str')
        
        xdf = xdf.dropna(axis='index', how='all').dropna(axis='columns', how='all')
        xdf.columns = ["col" +str(i)  for i in range(len(xdf.columns))]
        posicao = xdf.replace(np.nan, '').to_string().upper()
        re_mesref = _regex_mesref.search(posicao)
        mesref = 'N/D'
        if re_mesref:
            mesref = re_mesref.group(1)
            mesref = mesref[-4:]+'-'+mesref[3:5]
        if mesref == 'N/D': raise ValueError('Não foi possível encontrar o mês de referência no arquivo ' + i.as_posix())

        inicio = posicao.find("Renda Fixa".upper())
        fim = posicao.find("Proventos de Renda Fixa".upper())
        renda_fixa = posicao[inicio:fim].replace('FLU ','')
        
        
        p_valor = r'\d+(?:\.\d+)?'
        r = re.compile(r'(?P<tipo>LCA|LCI|LC|CDB)\s(\w+)\s+([\s|\w]+).+?(\d{4}-\d{2}-\d{2})\s00:00:00\s+('+p_valor+')\s+(\d{4}-\d{2}-\d{2})\s00:00:00\s+('+p_valor+')\s+('+p_valor+')\s+('+p_valor+')\s+('+p_valor+')')
        lancamentos = r.findall(renda_fixa)
        lancamentos = [list(i)+[mesref, arquivo] for i in lancamentos]
        lista +=  lancamentos
        #len(lista_renda_fixa) == 9 # esperado 9 items
        
    
    df = pd.DataFrame(lista, 
        columns = ['Tipo', 'Codigo', 'Nome', 'Vencimento', 'ValorPU', 'Data', 'ValorBruto', 'IR', 'IOF', 'ValorLiquido', 'AnoMes', 'Arquivo'],
        )
    for i in ['ValorPU', 'ValorBruto', 'ValorLiquido']:
        df[i] = df[i].astype(float)
    for i in ['Vencimento', 'Data']:
        df[i] = pd.to_datetime(df[i])

    return df

if __name__ == '__main__':

    compila_posicao()
