import json
import sqlite3 as sq
from unicodedata import normalize
import datetime
import csv
import os
import Bancos as bds
from os import listdir, getcwd, getpid
from os.path import isfile, join
import io
import win32com.client as win32

def atualiza_excel(nome_arquivo):
    xl = win32.gencache.EnsureDispatch("Excel.Application")
    xl.Visible = True
    xl.DisplayAlerts = False
    wb = xl.Workbooks.Open(nome_arquivo)
    wb.RefreshAll()
    wb.Save()
    wb.Close()
    xl.DisplayAlerts = False
    xl.Visible = False
    xl.Quit()
    os.system('taskkill /f /IM EXCEL.EXE')
    os.system('taskkill /f /IM EXCEL.EXE')

def converte_tipos(objeto):
    if objeto is None:
        return 'Null'
    if isinstance(objeto,int):
        return str(objeto)
    if  isinstance(objeto, str):
        return normalize('NFKD', objeto).encode('ASCII', 'ignore').decode('ASCII').upper()
    if isinstance(objeto, datetime.date):
        return objeto.strftime("%Y-%m-%d")
    else:
        return str(objeto)

arquivo_consulta = """SELECT o.numero as numero_chamado,
       o.problema as id_probema,
       p.problema,
      -- o.descricao, 
       o.local as id_local,
       l.local as local,
       o.sistema as id_area,
       s.sistema as area,
       u.nome as atendente,
       a.nome as quem_abriu,
       o.data_abertura,
       o.data_fechamento,
       o.data_atendimento,
       (select GROUP_CONCAT(p.num_pa SEPARATOR ', ') from pa_ocorrencias po join pa as p on p.id_pa = po.id_pa where po.id_ocorrencia = o.numero) as pas,
       TIMESTAMPDIFF(second,o.data_abertura,o.data_fechamento)/60 as tempo_gasto_min,
       CASE WHEN (SELECT DATABASE() FROM DUAL)='ocomon_novo' then 'POA' else 'SP' end as site
     --  ,(select assentamento from assentamentos where ocorrencia = o.numero order by numero desc limit 1) as solucao
FROM ocorrencias as o 
       LEFT JOIN sistemas as s on o.sistema = s.sis_id
       LEFT JOIN problemas as p on o.problema = p.prob_id
       LEFT JOIN localizacao as l on o.local = l.loc_id
       LEFT JOIN usuarios as u on o.operador = u.user_id
       LEFT JOIN (SELECT DISTINCT u.user_id, u.nome FROM usuarios as u INNER JOIN ocorrencias as oc on u.user_id = oc.aberto_por) as a on o.aberto_por = a.user_id
WHERE o.data_abertura >= ('2019-03-01') and s.sistema like 'TI %'"""

def executar_consulta(arquivo_consulta,string_conexao_origem,tipo_banco_origem):
    sql  = arquivo_consulta
    cursor_origem = bds.conectaBancos(string_conexao_origem,tipo_banco_origem)
    print("Vou executar a consulta")      
    cursor_origem.execute(sql)
    print("Vou exportar")
    field_names = [i[0] for i in cursor_origem.description]        
    resultados = [list(x) for x in cursor_origem.fetchall()]
    resultados_convertidos = []
    for i in resultados:
        linha = []
        for j in i:
            x = converte_tipos(j)
            linha.append(x)
        resultados_convertidos.append(linha)
    return resultados_convertidos  

conexao_ocomon_sp = "172.20.0.150,ocomon_integrador,root,auditek"
conexao_ocomon_poa = "192.168.255.151,ocomon_novo,root,xparioocomon"

resultado_sp = executar_consulta(arquivo_consulta,conexao_ocomon_sp,'mysql')
resultado_poa = executar_consulta(arquivo_consulta,conexao_ocomon_poa,'mysql')



field_names = ['numero_chamado',	'id_probema',	'problema',	'id_local',	'local',	'id_area',	'area',	'atendente',	'quem_abriu',	'data_abertura',	'data_fechamento',	'data_atendimento',	'pas',	'tempo_gasto_min',	'site','classe']
resultado_final = resultado_sp + resultado_poa

classificacoes = json.loads(open('problemas_classificados.json').read())

print(classificacoes.keys())

exit()
with open('saida_final.csv', 'w') as f:
        writer = csv.writer(f,delimiter=";")
        writer.writerow(field_names)
        for i in resultado_final:
            writer.writerow(i)

atualiza_excel(os.getcwd()+"\\"+"analise_ocomon.xlsx")