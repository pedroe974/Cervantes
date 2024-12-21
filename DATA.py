from datetime import datetime,timedelta
from dateutil.relativedelta import relativedelta
import numpy as np
import pandas as pd

# Obtendo o mês e o ano atual
hoje = pd.Timestamp.now()
inicio_do_mes = hoje.replace(day=1)
fim_do_mes = hoje + pd.offsets.MonthEnd()

# Gerando uma lista de datas do mês atual
datas_do_mes = pd.date_range(start=inicio_do_mes, end=fim_do_mes)

# Contando os dias úteis (excluindo fins de semana)
dias_uteis = np.busday_count(inicio_do_mes.date(), fim_do_mes.date() + pd.Timedelta(days=1))
dias_uteis=str(dias_uteis)


# Obter o ano atual
hoje = datetime.today()
ano_atual = datetime.now().year

# Primeiro dia do mesmo mês do ano anterior
primeiro_dia_mes_ano_anterior = hoje.replace(year=hoje.year - 1, month=hoje.month, day=1)

# Último dia do mesmo mês do ano anterior
ultimo_dia_mes_ano_anterior = (primeiro_dia_mes_ano_anterior.replace(month=primeiro_dia_mes_ano_anterior.month + 1) - timedelta(days=1)
                               if primeiro_dia_mes_ano_anterior.month < 12
                               else primeiro_dia_mes_ano_anterior.replace(day=31))

# Primeiro dia do mês anterior
primeiro_dia_mes_anterior = (hoje.replace(day=1) - timedelta(days=1)).replace(day=1)

# Último dia do mês anterior
ultimo_dia_mes_anterior = hoje.replace(day=1) - timedelta(days=1)

# Primeiro dia do ano
primeiro_dia_do_ano = datetime(ano_atual, 1, 1)
primeiro_dia_do_ano=primeiro_dia_do_ano.strftime("%d/%m/%Y")
# Primeira e última data dos próximos 30 dias
primeira_data = (hoje + timedelta(days=1)).strftime('%d/%m/%Y')
ultima_data = (hoje + timedelta(days=30)).strftime('%d/%m/%Y')
# Obter a data atual
hoje = datetime.today()

# Verificar se o dia atual é o primeiro dia do mês
if hoje.day == 1:
    # Se for o primeiro dia, calcular o primeiro dia do mês anterior
    data_inicial = (hoje - relativedelta(months=1)).replace(day=1)
    # Último dia do mês anterior
    ultimo_dia = hoje.replace(day=1) - relativedelta(days=1)
    # Dia anterior (do último mês)
    dia_anterior = ultimo_dia
else:
    # Primeiro dia do mês atual
    data_inicial = hoje.replace(day=1)
    # Último dia do mês atual
    ultimo_dia = (hoje + relativedelta(months=1)).replace(day=1) - relativedelta(days=1)
    # Dia anterior (do dia atual)
    dia_anterior = hoje - relativedelta(days=1)

dia=hoje.strftime("%d")
data_inicialf=data_inicial.strftime("%d%m%Y")
dia_anteriorf=dia_anterior.strftime("%d%m%Y")
data_inicial=data_inicial.strftime("%d/%m/%Y")
dia_anterior=dia_anterior.strftime("%d/%m/%Y")
mes=ultimo_dia.strftime("%m")
ano=ultimo_dia.strftime("%Y")

primeiro_dia_mes_anterior = primeiro_dia_mes_anterior .strftime("%d/%m/%Y")
ultimo_dia_mes_anterior = ultimo_dia_mes_anterior .strftime("%d/%m/%Y")
primeiro_dia_mes_ano_anterior=primeiro_dia_mes_ano_anterior.strftime("%d/%m/%Y")
ultimo_dia_mes_ano_anterior = ultimo_dia_mes_ano_anterior.strftime("%d/%m/%Y")
# Subtraindo 3 meses
tres_meses_antes = hoje - relativedelta(months=3)
tres_mes_atras=tres_meses_antes.strftime("%m")
# Obtendo o primeiro dia do mês
primeiro_dia_tres_meses_antes = tres_meses_antes.replace(day=1)
primeiro_dia_tres_meses_antes=primeiro_dia_tres_meses_antes.strftime("%d/%m/%Y")