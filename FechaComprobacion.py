from datetime import datetime, timedelta
import calendar

#def esFDM ():
#    hoy = date.today()

#    ultimo_dia = calendar.monthrange(hoy.year, hoy.month)[1]

#    es_fin_de_mes = hoy.day == ultimo_dia

#    return es_fin_de_mes

def esIDM():
    hoy = datetime.now()

    es_primer_dia = hoy.day == 1

    return es_primer_dia


def diasMesPasado ():

    hoy = datetime.now()

    primer_dia_actual = hoy.replace(day=1)
    un_dia_del_mes_pasado = primer_dia_actual - timedelta(days=1)

    primer_dia_pasado = un_dia_del_mes_pasado.replace(day=1)

    _, total_dias_mes_pasado = calendar.monthrange(un_dia_del_mes_pasado.year, un_dia_del_mes_pasado.month)
    ultimo_dia_pasado = un_dia_del_mes_pasado.replace(day=total_dias_mes_pasado)

    primer_dia = primer_dia_pasado.strftime("%m/%d/%Y")
    ultimo_dia = ultimo_dia_pasado.strftime("%m/%d/%Y")

    return primer_dia, ultimo_dia

