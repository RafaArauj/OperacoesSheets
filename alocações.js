///////

//função para puxar todos os eventos dos Sementers na semana passada
function AlocaçõesSemente() {

    //registra o tempo de início da função [ainda não utilizado]
    //var start = new Date().UTC();
    
    //chama a aba Alocações da planilha pelo ID
    var ss = SpreadsheetApp.openById('1zKhN7YI5XIRXpGBd1CVrY3Cr5TQGXUiufsNHi7MqU5M');
    var sheets = ss.getSheets();
    var sheet = ss.getSheetByName('Alocações 2');
    //inserir linha depois da linha 3 pra manter as formulas funcionando
    //sheet.insertRowBefore(4);
    
    //chama a aba Salários da planilha pelo ID e puxa os valores de Sementers[0] e Emails[1]
    var reference = ss.getSheetByName('Sementers');
    var sementers = reference.getRange(2,2,reference.getLastRow(),3).getValues();
    
    //define o período a ser consultado, com a data atual, ano e mês atual, sete dias atrás
    var now = new Date();
    var year = now.getFullYear();
    var month = now.getMonth();
    var day = now.getDate()-1;
    //se quiser rodar o programa até outro dia que não ontem
    //var day = now.getDate()-2;
    var week = day - 7;
    var date = now.toString();
    //se quiser rodar o programa para o mês inteiro ou para períodos diferentes
    //var week = 1;
    //last.setValue(Date);
    
    /////////////para cada sementer/////////////
    for (var i = 0; i < sementers.length; i++){
    //se quiser rodar começando de outro Sementer
    //for (var i = 31; i < sementers.length; i++){
    ////////////////////////////////////////////
      
      //nome do sementer
      var sementer = sementers[i][0];
      Logger.log(sementer);
      //email do sementer
      var email = sementers[i][1];
      Logger.log(email);
      
      //puxa o calendário do sementer
      //é preciso que o usuário do Google que roda a função tenha se inscrito na agenda a ser consultada
      var calendar = CalendarApp.getCalendarById(email);
      
      //Verifica se a agenda (email) ainda existe e se o usuário que rodou o programa está inscrito na agenda
      if (calendar !== null){
        
        //puxa todos os eventos do período definido anteriormente
        var events = calendar.getEvents(new Date(year, month, week, 00, 00, 00, 00), new Date(year, month, day, 00, 00, 00, 00));
  
        //para cada evento
        for (var j = 0; j < events.length; j++){
          var event = events[j];
  
          //verifica se o evento é de dia inteiro e 
          //[AINDA NÃO IMPLEMENTADO]se foi confirmado pelo sementer && event.email.getGuestStatus()!="NO" && email.getGuestStatus()!="INVITED" && email.getGuestStatus()!="MAYBE"
          if (event.isAllDayEvent()==false) {
          
            //título do evento
            var title = event.getTitle();
  
            //verifica se é um evento da Semente. 
            //Aqui usamos um código de tags entre colchetes [TAG]. Eventos sem tag são considerados eventos pessoais
            if (title.indexOf("[") >= 0 && title.indexOf("]") >= 0){
  
              //[AINDA NÃO IMPLEMENTADO - melhoria de performance]
              //computar o número de eventos com tag
              //criar um range do mesmo tamanho
              //colocar os dados importantes desses eventos no range
              //"pushar" o range inteiro, de uma só vez, para a planilha
  
              //nome do sementer na coluna A
              var c = 1
              var dataRange = sheet.getRange(4,c);
              dataRange.setValue(sementer);
              c++
            
              //dia do início do evento na coluna B
              var date = event.getStartTime();
              dataRange = sheet.getRange(4,c);
              dataRange.setValue(date);
              c++
  
              //string ente [] do título do evento na coluna C
              var project = title.substring(title.lastIndexOf("[")+1,title.lastIndexOf("]"));
              dataRange = sheet.getRange(4,c);
              dataRange.setValue(project);
              c++
  
              //calcula a duração do evento, em horas na coluna D
              var duration = (event.getEndTime()-event.getStartTime())/3600000;
              dataRange = sheet.getRange(4,c);
              dataRange.setValue(duration);
  
              //inserir linha depois da linha 3 pra manter as formulas da linha de referência funcionando
              sheet.insertRowBefore(4);
            }
          }
        }
      }
      //caso não encontra a agenda, passa para o próximo Sementer 
      //ainda não dá uma mensagem de erro, ou marca o Sementer com um indicador
      else {
        i++
      }
    }
    
    //preenche automaticamente as colunas E a J com as formulas de processamento daqueles dados
    var column = 5;
    var reference = sheet.getRange(3, column, 1, 6);
    reference.autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    
    //registra o tempo de encerramento da execução
    //var last = sheet.getRange(1,10);
    //var end = new Date().UTC();
    //var time = end - start;
    //last.setValue(time);
      //Logger.log(time);  
  }
  
  ///////
  
  //função para salvar um registro dos salários no mês que a função AlocaçãoSemente foi rodada
  function SalvarSalários() {
    
    //chama a aba Salários da planilha pelo ID
    var ss = SpreadsheetApp.openById('1zKhN7YI5XIRXpGBd1CVrY3Cr5TQGXUiufsNHi7MqU5M');
    var sheets = ss.getSheets();
    var reference = ss.getSheetByName('Sementers');
    
    //puxa os valores de Salários
    var salaries = reference.getRange(1,5,reference.getLastRow(),1).getValues();
    
    //cria uma matriz com os meses
    var months = new Array();
    months[0] = "janeiro";
    months[1] = "fevereiro";
    months[2] = "março";
    months[3] = "abril";
    months[4] = "maio";
    months[5] = "junho";
    months[6] = "julho";
    months[7] = "agosto";
    months[8] = "setembro";
    months[9] = "outubror";
    months[10] = "novembro";
    months[11] = "dezembro";
    
    //atribui à n o número relativo ao mês passado
    var now = new Date();
    var year = now.getFullYear();
    var n = now.getMonth()-1;
    
    //gambiarra pq -1.0 não é = 11
    if (n == 0){
      n = 11
    }
    
    //insere uma coluna após a coluna de salários atuais
    reference.insertColumnAfter(5);
    
    //copia os valores dos salários para a coluna D
    var dataRange = reference.getRange(1,6,reference.getLastRow(),1);
    dataRange.setValues(salaries);
    
    //estabelece o nome do mês ao valor da célula D1
    var month = months[n];
    dataRange = reference.getRange(1,6);
    dataRange.setValue(month+"/"+year);
    }
  
  ///////
  
  //função para salvar um registro dos salários no mês que a função AlocaçãoSemente foi rodada
  function SalvarAlocações() {
    
    //chama a aba Salários da planilha pelo ID
    var ss = SpreadsheetApp.openById('1zKhN7YI5XIRXpGBd1CVrY3Cr5TQGXUiufsNHi7MqU5M');
    var sheets = ss.getSheets();
    var reference = ss.getSheetByName('Projetos');
    
    //puxa os valores de Salários
    var salaries = reference.getRange(1,5,reference.getLastRow(),1).getValues();
    
    //cria uma matriz com os meses
    var months = new Array();
    months[0] = "janeiro";
    months[1] = "fevereiro";
    months[2] = "março";
    months[3] = "abril";
    months[4] = "maio";
    months[5] = "junho";
    months[6] = "julho";
    months[7] = "agosto";
    months[8] = "setembro";
    months[9] = "outubror";
    months[10] = "novembro";
    months[11] = "dezembro";
    
    //atribui à n o número relativo ao mês passado
    var now = new Date();
    var year = now.getFullYear();
    var n = now.getMonth()-1;
    
    //gambiarra pq -1.0 não é = 11
    if (n == 0){
      n = 11;
    }
    
    //insere uma coluna após a coluna de salários atuais
    reference.insertColumnAfter(5);
    
    //copia os valores dos salários para a coluna D
    var sourceRange = reference.getRange(1,7,reference.getLastRow(),1);
    var dataRange = reference.getRange(1,6,reference.getLastRow(),1);
    sourceRange.autoFill(dataRange, destination, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    
    //estabelece o nome do mês ao valor da célula D1
    var month = months[n];
    dataRange = reference.getRange(1,6);
    dataRange.setValue(month+"/"+year);
    }
  
  ///////
  
  //função para mandar email com o relatório de alocação mensal
  function EmailAlocação() {
    //chama a aba Cálculos Dashboard Ano da planilha pelo ID
    var ss = SpreadsheetApp.openById('1zKhN7YI5XIRXpGBd1CVrY3Cr5TQGXUiufsNHi7MqU5M');
    var sheets = ss.getSheets();
    var sheet = ss.getSheetByName('Cálculos Dashboard Ano');
    
    var agora = new Date();
    var mes = agora.getMonth();
    var mescoluna = mes+1;
    
    var verticais = sheet.getRange('A3:A5').getValues();
    Logger.log(verticais);
    var taxa = sheet.getRange('R3C'+mescoluna+':R6C'+mescoluna).getValues();
    
    var emails = 'ellen@sementenegocios.com.br, alline@sementenegocios.com.br, pablo@sementenegocios.com.br, marcel@sementenegocios.com.br, marcio@sementenegocios.com.br, cesar@sementenegocios.com.br'
    var assunto = 'Relatório mensal de alocação';
    var corpo = 'Olá! Este é um email automático, mas responder ele funciona.\n\nA seguir, as taxas de alocação em projetos externos para cada vertical para o mês passado ('+mes+'):\n\n'+verticais[0]+': '+taxa[0]+'%\n'+verticais[1]+': '+taxa[1]+'%\n'+verticais[2]+': '+taxa[2]+'%\n\nEste relatório é enviado todo dia 8 de cada mês, reportando o resultado do mês passado. O delay de 8 dias é necessário para que todas as alocações sejam computadas.\n\n*Esse valor é igual a TAE/DOV, onde TAE é o total de horas alocadas em projetos externos por todos da vertical (incluindo heads) no mês, e DOV é a disponibilidade ótima para aquela vertical, em horas.\n**A disponibilidade ótima para cada vertical DOV é calculada multiplicando-se o número de consultores que aquela vertical conta no mês (C), por 70% das horas disponíveis de trabalho naquele mês (H).\n***Horas disponíveis (H) correspondem ao total de dias úteis multiplicado por 8 horas diárias.';
    
    MailApp.sendEmail({
      to: emails,
      subject: assunto,
      body: corpo, 
    })
    }
  
  //função teste de funcionalidades
  function testeobjeto() {
    var zeh = 'jose@sementenegocios.com.br';
    var calendario = CalendarApp.getCalendarById(zeh);
    var eventos = calendario.getEvents(new Date(2019, 03, 11, 00, 00, 00, 00), new Date(2019, 04, 12, 00, 00, 00, 00));
    
    for(var i = 0; i<eventos.length; i++){
      var evento = eventos[i]
    }
  }