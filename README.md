# relatorioAutomatico

Envio automático de relatório semanal via email

Bibliotecas utilizadas: 
-os 
-selenium
-win32com.client 
-pythoncom 
-time 
-datetime

Funcionalidades:
-Acessa o sistema da empresa, inserindo login e senha;

-Identifica a data atual e com base na mesma retorna as datas referentes a segunda feira da semana anterior e ao domingo da semana atual;

-Busca o relatório de equipes referente ao intervalo de tempo entre as datas obtidas;

-Baixa o relatório;

-Identifica o horário atual e de acordo com o mesmo define se a saudação do email será "Bom dia!", "Boa tarde!" ou "Boa noite!"

-Envia o relatório via email através de integração com o Outlook

-Deleta o relatório baixado após a confirmação do envio de email
