import pyodbc
import MySQLdb
import psycopg2

def conectaBancos(dadosConexao, tipoBanco):
    bd = dadosConexao.split(",")
    if(tipoBanco=='postgres'):
        dadosConexao = "host='"+bd[0]+"' dbname='"+bd[1]+"' user='"+bd[2]+"' password='"+bd[3]+"'"
        db=psycopg2.connect(dadosConexao)
    elif tipoBanco == 'mysql':
        db= MySQLdb.connect(host=bd[0], user=bd[2], passwd=bd[3], db=bd[1])
    elif tipoBanco == 'sql_server':
        db = pyodbc.connect(r'DRIVER={SQL Server};'r'SERVER=' + bd[0] + ';'r'DATABASE=' + bd[1] + ';'r'UID=' + bd[2] + ';'r'PWD=' + bd[3] + '')
    cursor = db.cursor()
    return cursor