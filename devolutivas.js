// Função para criar um menu "Exportar" contendo as funções de impressão e envio de email
function onOpen() {
  
    //
    var submenu = [
      {name:"Salvar devolutiva como PDF", functionName:"gerarPDF"},
      {name:"Salvar todas as devolutivas como PDF", functionName:"gerarPDFs"},
      {name:"Enviar devolutiva por email", functionName:"mandarEmail"},
      {name:"Enviar todas as devolutivas por email", functionName:"mandarEmails"}
    ];
    SpreadsheetApp.getActiveSpreadsheet().addMenu('Exportar', submenu);  
  }
  
  ///////
  
  // Função para imprimir a devolutiva que está sendo mostrada e salvá-la como PDF da pasta de Devolutivas
  function gerarPDF() {
    
    // Chama a planilha do Delas
    var fontePlanilha = SpreadsheetApp.openById('1BOpUCp3_oCTqvL3Lr73pLysS44j5x6g8d0rG4Cz-imA');
    
    // Chama a aba de Devolutiva
    var planilha = fontePlanilha.getSheetByName('Devolutiva');
    
    // Define o nome do arquivo como nome da empreendedora
    var nome = planilha.getRange(47,4).getValue();
    var nomePDF = nome;
    SpreadsheetApp.flush();
  
    // Chama a pasta Devolutivas
    var foldersave = DriveApp.getFolderById('1Moo0V5JwMomr6fSvS21W4Z-4Jz2b4NQL');
    
    var request = {
      "method": "GET",
      "headers":{"Authorization": "Bearer "+ScriptApp.getOAuthToken()},    
      "muteHttpExceptions": true
    };
  
    // 
    var fetch = 'https://docs.google.com/spreadsheets/d/1BOpUCp3_oCTqvL3Lr73pLysS44j5x6g8d0rG4Cz-imA/export?gid=389510879&range=A1:AI51&format=pdf&size=A4&portrait=true&scale=1&&top_margin=0.00&bottom_margin=0.00&left_margin=0.00&right_margin=0.00'
  
    // 
    var pdf = UrlFetchApp.fetch(fetch, request);
    pdf = pdf.getBlob().getAs('application/pdf').setName(nomePDF);
    var file = foldersave.createFile(pdf)
  }
  
  ///////
  
  function gerarPDFs() {
    // Chama a planilha do Delas
    var sourceSpreadsheet = SpreadsheetApp.openById('1BOpUCp3_oCTqvL3Lr73pLysS44j5x6g8d0rG4Cz-imA');
    
    // Chama as abas de Respotas e de Devolutiva
    var sheets = sourceSpreadsheet.getSheets();
    var respostas = sourceSpreadsheet.getSheetByName('Valores');
    var sourceSheet = sourceSpreadsheet.getSheetByName('Devolutiva');
  
    // Chama a pasta do Drive na qual vão ser salvas as Devolutivas
    var foldersave = DriveApp.getFolderById('1Moo0V5JwMomr6fSvS21W4Z-4Jz2b4NQL');
    
    var request = {
      "method": "GET",
      "headers":{"Authorization": "Bearer "+ScriptApp.getOAuthToken()},    
      "muteHttpExceptions": true
    };
  
    var fetch='https://docs.google.com/spreadsheets/d/1BOpUCp3_oCTqvL3Lr73pLysS44j5x6g8d0rG4Cz-imA/export?gid=389510879&range=A1:AI51&format=pdf&size=A4&portrait=true&scale=1&&top_margin=0.00&bottom_margin=0.00&left_margin=0.00&right_margin=0.00'
    
    var empreendedoras = respostas.getRange(2,2,respostas.getLastRow()-1,1).getValues();
    
    //for (var i = 0; i < empreendedoras.length; i++){
    for (var i = 782; i < empreendedoras.length; i++){
      var nome = sourceSheet.getRange(45,3);
      nome.setValue(empreendedoras[i]);
  
      SpreadsheetApp.flush();
  
      var pdf = UrlFetchApp.fetch(fetch, request);
      pdf = pdf.getBlob().getAs('application/pdf').setName(empreendedoras[i]);
      var file = foldersave.createFile(pdf);
      Utilities.sleep(9000);
      
      Logger.log(i);
      Logger.log(empreendedoras[i]);
    }
  }
  
  ///////
  
  function mandarEmail(){
    // Chama a planilha do Delas
    var fontePlanilha = SpreadsheetApp.openById('1BOpUCp3_oCTqvL3Lr73pLysS44j5x6g8d0rG4Cz-imA');
    
    // Chama as abas de Respotas e de Devolutiva
    var planilha = fontePlanilha.getSheetByName('Devolutiva');
    var grafico = fontePlanilha.getSheetByName('Gráfico');
    
    // Chama a pasta do Drive com as Devolutivas
    var folderget = DriveApp.getFolderById('1Moo0V5JwMomr6fSvS21W4Z-4Jz2b4NQL');
    gerarPDF();
    
    var nome = grafico.getRange(1,2).getValue();
    var trilha = grafico.getRange(3,2).getValue();
    var email = 'planilha.getRange(48,4).getValue()';
    var nomearquivo = planilha.getRange(47,4).getValue();
    var arquivo = DriveApp.getFilesByName(nomearquivo);
    var assunto = 'Aqui está o seu diagnóstico do Delas!';
    var corpo = 'Olá '  + nome + ', tudo bem? \n\nParabéns! Você foi selecionada para a turma 2021 do Programa de Desenvolvimento Sebrae Delas Mulher de Negócios. \n\nTeremos um ano de muito conteúdo e troca entre empreendedoras de toda Santa Catarina. \n\nNossos encontros acontecerão pela plataforma StudiOn (https://studion-sebraedelassc.dotgroup.com.br/auth/signin). Fique tranquila que em breve você receberá um e-mail com as orientações de acesso. \n\nNosso encontro de lançamento será no dia 15 de abril, das 9h às 11h30. Neste dia, vamos apresentar o programa, passar todas as informações necessárias, tirar possíveis dúvidas e receber uma convidada especial que abordará o tema "Competências da Empreendedora do Futuro". Você poderá acessar o evento na plataforma StudiOn utilizando seu login e senha. \n\nOs encontros de conteúdo começarão no dia 29 de abril e acontecerão sempre às quintas-feiras, das 9h às 11h30, a cada dia 15 dias. Recomendamos já bloquear sua agenda. \n\nToda a comunicação do programa será feita por meio do StudioOn, mas também teremos um grupo fechado no WhatsApp. Para entrar nele basta clicar aqui: https://chat.whatsapp.com/BcD22jsswFwI0if4f8JVWp. \n\nAlém disso, caso você tenha qualquer dúvida e queira entrar em contato com a equipe do Programa, nosso e-mail oficial é sebraedelas@sc.sebrae.com.br. \n\nPara começar essa jornada, que tal olhar para dentro? Você está recebendo aqui pela Semente Negócios, parceira do Sebrae SC que nos acompanhará neste programa, o diagnóstico do seu negócio a partir das suas respostas ao formulário de inscrição. Ele é um olhar cuidadoso neste momento de tantas mudanças e servirá como guia para trabalharmos juntas os necessários pontos de evolução. \n\nContamos com o seu comprometimento para o sucesso desta iniciativa. \n\nVamos juntas? \n\nEquipe Sebrae Delas Mulher de Negócios \nSanta Catarina \n#sebraedelassc \n#juntassomosmaisfortes \n@sebraesc';
  
    MailApp.sendEmail({
      to: email,
      subject: assunto,
      body: corpo, 
      attachments: [arquivo.next().getAs(MimeType.PDF)],
      bcc: 'alline@sementenegocios.com.br,jose@sementenegocios.com.br,manoela@sementenegocios.com.br',
    })
  
    // Marcar devolutivas enviadas
  }
  
  ///////
  
  function mandarEmails(){
    // Chama a planilha do Delas
    var fontePlanilha = SpreadsheetApp.openById('1BOpUCp3_oCTqvL3Lr73pLysS44j5x6g8d0rG4Cz-imA');
    
    // Chama as abas de Respotas e de Devolutiva
    var valores = fontePlanilha.getSheetByName('Valores');
    //var respostas = fontePlanilha.getSheetByName('Respostas ao formulário 5');
    
    // Chama a pasta do Drive com as Devolutivas
    var folderget = DriveApp.getFolderById('1Moo0V5JwMomr6fSvS21W4Z-4Jz2b4NQL');
    //var folderget2 = DriveApp.getFolderById('');
    //var folderget3 = DriveApp.getFolderById('');
    
    var empreendedoras = valores.getRange(2,2,valores.getLastRow()-1,1).getValues();
    //var empreendedorasMinusculo = respostas.getRange(2,2,respostas.getLastRow()-1, 1).getValues();
    var emails = valores.getRange(2,4,valores.getLastRow()-1,1).getValues();
    Logger.log(empreendedoras.length)
  
    for (var i = 0; i < 1; i++){
    //for (var i = 0; i < empreendedoras.length; i++){
      var nome = empreendedoras[i];
      //var nomeMinusculo = empreendedorasMinusculo[i];
      var arquivo = folderget.getFilesByName(nome);
      //var arquivo2 = folderget2.getFilesByName(nome);
      //var arquivo3 = folderget3.getFilesByName(nomeMinusculo+'.pdf');
      var vemail = emails[i];
      var email = vemail.toString();
      var assunto = 'Aqui está o seu diagnóstico do Delas!';
      var corpo = 'Olá '  + nome + ', tudo bem? \n\nParabéns! Você foi selecionada para a turma 2021 do Programa de Desenvolvimento Sebrae Delas Mulher de Negócios. \n\nTeremos um ano de muito conteúdo e troca entre empreendedoras de toda Santa Catarina. \n\nNossos encontros acontecerão pela plataforma StudiOn (https://studion-sebraedelassc.dotgroup.com.br/auth/signin). Fique tranquila que em breve você receberá um e-mail com as orientações de acesso. \n\nNosso encontro de lançamento será no dia 15 de abril, das 9h às 11h30. Neste dia, vamos apresentar o programa, passar todas as informações necessárias, tirar possíveis dúvidas e receber uma convidada especial que abordará o tema "Competências da Empreendedora do Futuro". Você poderá acessar o evento na plataforma StudiOn utilizando seu login e senha. \n\nOs encontros de conteúdo começarão no dia 29 de abril e acontecerão sempre às quintas-feiras, das 9h às 11h30, a cada dia 15 dias. Recomendamos já bloquear sua agenda. \n\nToda a comunicação do programa será feita por meio do StudioOn, mas também teremos um grupo fechado no WhatsApp. Para entrar nele basta clicar aqui: https://chat.whatsapp.com/BcD22jsswFwI0if4f8JVWp. \n\nAlém disso, caso você tenha qualquer dúvida e queira entrar em contato com a equipe do Programa, nosso e-mail oficial é sebraedelas@sc.sebrae.com.br. \n\nPara começar essa jornada, que tal olhar para dentro? Você está recebendo aqui pela Semente Negócios, parceira do Sebrae SC que nos acompanhará neste programa, o diagnóstico do seu negócio a partir das suas respostas ao formulário de inscrição. Ele é um olhar cuidadoso neste momento de tantas mudanças e servirá como guia para trabalharmos juntas os necessários pontos de evolução. \n\nContamos com o seu comprometimento para o sucesso desta iniciativa. \n\nVamos juntas? \n\nEquipe Sebrae Delas Mulher de Negócios \nSanta Catarina \n#sebraedelassc \n#juntassomosmaisfortes \n@sebraesc';
      Logger.log(email)
      MailApp.sendEmail({
        to: email,
        subject: assunto,
        body: corpo, 
        attachments: [arquivo.next().getAs(MimeType.PDF)/*,
                      arquivo2.next().getAs(MimeType.PDF),
                      arquivo3.next().getAs(MimeType.PDF)*/],
        //bcc: 'alline@sementenegocios.com.br',
      })
      Logger.log(i);
    }
    
    // Marcar devolutivas enviadas
  }
  
  