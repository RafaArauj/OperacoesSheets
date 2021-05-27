//função para puxar todos os deals do pipedrive
function GetPipedriveDeals() {

    //chama e limpa a aba Automação da planilha atual
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheets = ss.getSheets();
      var sheet = ss.getSheetByName('Automação');
      var cells = sheet.getRange(2,1,sheet.getLastRow(),sheet.getLastColumn());
      cells.clear();
      
    //constrói uma página do url JSON com os deals do Pipedrive
      var firstdeal = 0
      var url1 = "https://api.pipedrive.com/v1/deals?all_not_deleted&start="
      var start = firstdeal;
      var url2 = "&limit=";
      var limit = 500;
      var url3  = "&api_token=";
      var token = "6c309115ea9eaea89650b14ad8186921008f0e81" //token de API do Pipedrive da Sensorweb*
      var url = (url1+start+url2+limit+url3+token);
    
    //chama a API do Pipedrive e preenche "dataAll" com JSON.parse(), então a informação de "data" é atribuída ao "dataSet" para que se possa preencher "dataAll" com novos dados
      var response = UrlFetchApp.fetch(url);
      var dataAll = JSON.parse(response.getContentText());
      var dataSet = dataAll["data"];
      
    //loop para puxar todas as páginas
      do {
      
    //prepara as variáveis para o loop de preenchimento da planilha
      var rows = []; //cria a matriz na qual os dados vão ser colocados
      var page = sheet.getLastRow()+1;
      var row = sheet.getLastRow()+1;
      var range = 16 //o número total de campos que cada objeto em JSON possui
      var data;
      var dataRange;
      
    //preenche a matriz com os dados do JSON
      for (var i = 0; i < dataSet.length; i++){
    
      //atribui a entrada "i" de "dataSet" à "data"
        data = dataSet[i];
        
      //empurra a linha "i" de dados na matriz "rows"
        rows.push([data.title,
                   data.id,
                   data.value,
                   data.weighted_value,
                   data.pipeline_id,
                   data.stage_id,
                   data.status,
                   data.add_time,
                   data.update_time,
                   data.next_activity_date,
                   data.last_activity_date,
                   data.won_time,
                   data.lost_time,
                   data.close_time,
                   data.lost_reason,
                   data.activities_count,
                   ]);
        
      //seleciona a série de células ao qual os valores de "rows" serão atribuídos
        dataRange = sheet.getRange(page, 1, rows.length, range);
          
      //atribui os valores de "rows" à série de células "dataRange"
        dataRange.setValues(rows);
            
      //passa a referência de célula a ser preenchida para a próxima linha
        row = row + 1
      }
    
    //prepara o próximo loop
      start = start + limit;
      var url = (url1+start+url2+limit+url3+token)
      var response = UrlFetchApp.fetch(url);
      var dataAll = JSON.parse(response.getContentText());
      var dataSet = dataAll["data"];
      }
      
    //quebra o loop quando não há mais dados
      while(dataSet !== null);
    }
    
    
    
    
    
    
    
    //função para puxar todas as organizações do pipedrive
    function GetPipedriveOrgs() {
    
    //chama e limpa a aba Organizações da planilha
      var ss = SpreadsheetApp.openById("1JhhOQ-bWi6Lh66RpBwftyJLk1dB-dd6GgeAziU3oJCA"); //abrir a planilha pelo URL*
      var sheets = ss.getSheets();
      var sheet = ss.getSheetByName('Organizações');
      
    //constrói uma página do url JSON com as Organizações do Pipedrive
      var url1 = "https://api.pipedrive.com/v1/organizations?start="
      var start = 0;
      var url2 = "&limit=";
      var limit = 500;
      var url3  = "&api_token=";
      var token = "6c309115ea9eaea89650b14ad8186921008f0e81" //token de API do Pipedrive da Sensorweb*
      var url = (url1+start+url2+limit+url3+token);
    
    //chama a API do Pipedrive e preenche "dataAll" com JSON.parse(), então a informação de "data" é atribuída ao "dataSet" para que se possa preencher "dataAll" com novos dados
      var response = UrlFetchApp.fetch(url);
      var dataAll = JSON.parse(response.getContentText());
      var dataSet = dataAll["data"];
      
    //loop para puxar todas as páginas
      do {
      
    //prepara as variáveis para o loop de preenchimento da planilha
      var rows = []; //cria a matriz na qual os dados vão ser colocados
      var page = sheet.getLastRow()+1;
      var row = sheet.getLastRow()+1;
      var range = 2 //o número total de campos que cada objeto em JSON possui
      var data;
      var dataRange;
      
    //preenche a matriz com os dados do JSON
      for (var i = 0; i < dataSet.length; i++){
    
      //atribui a entrada "i" de "dataSet" à "data"
        data = dataSet[i];
        
      //empurra a linha "i" de dados na matriz "rows"
        rows.push([data.id,data.org_id]);
        
      //seleciona a série de células ao qual os valores de "rows" serão atribuídos
        dataRange = sheet.getRange(page, 1, rows.length, range);
          
      //atribui os valores de "rows" à série de células "dataRange"
        dataRange.setValues(rows);
    
      //normaliza a entrada de dados dos códigos do setor adicionando 1s às células vazias
        var zero = sheet.getRange(row,range+1);
        zero.setFormula('=if(b'+row+'="";1;b'+row+')');
            
      //passa a referência de célula a ser preenchida para a próxima linha
        row = row + 1
    
      }
    
    //prepara o próximo loop
      start = start + limit;
      var url = (url1+start+url2+limit+url3+token)
      var response = UrlFetchApp.fetch(url);
      var dataAll = JSON.parse(response.getContentText());
      var dataSet = dataAll["data"];
      }
      
    //quebra o loop quando não há mais dados
      while(dataSet !== null);
    }
    
    