// @ts-nocheck
//Calendar

var app=SpreadsheetApp;
var calendario=CalendarApp.getCalendarById("c_u8g6cg8v7te69bnudr25qkr8l0@group.calendar.google.com");
var sheet=app.getActiveSheet();

//FUNÇÃO USADA PARA PEGAR OS DADOS DOS ANIVERSARIANTES E JOGAR NO GOOGLE AGENDAS

 function myCalendar()
 {

  var range=sheet.getRange("A3:B").getValues();
   range.map(function(elem,ind,obj){
     if(elem[0]!=""){

       calendario.createAllDayEvent(elem[0], elem[1]); 
     }   
   }); 
 }


//LINKAMENTO ENTRE GOOGLE SHEETS E GOOGLE AGENDAS

function cadastro()
{
var ss=SpreadsheetApp.getActiveSpreadsheet();
var nome_fun = ss.getRange('C6').getValue();
var telefone = ss.getRange('C9').getValue();
var data_aniv = ss.getRange('C12').getValue();
var cpf = ss.getRange('C15').getValue();
var funcao = ss.getRange('C18').getValue();
var data_nascimento = ss.getRange('C21').getValue();
var event;
var eventId;

if(nome_fun == "" || telefone == "" || data_aniv == "" || funcao == "" || data_nascimento == ""){
  SpreadsheetApp.getUi().alert("Insira todos os dados Obrigatorios!!");
  return;
}
event=calendario.createAllDayEvent(nome_fun,data_aniv);
eventId = event.getId();
var info = [nome_fun, telefone, data_aniv, cpf, funcao, data_nascimento, eventId]

ss.getRange('C6').clearContent();
ss.getRange('C9').clearContent();
ss.getRange('C12').clearContent();
ss.getRange('C15').clearContent();
ss.getRange('C18').clearContent();
ss.getRange('C21').clearContent();

var ss = ss.getSheetByName('Colaboradores 👨‍🏭');

ultima_linha = ss.getLastRow();

for (let i=0; i<7 ; i++){
  ss.getRange(ultima_linha+1,i+1).setValue(info[i]);
}
  SpreadsheetApp.getUi().alert("Cadastro Realizado!!");
}


//FUNÇÃO PARA BUSCAR COLABORADOR JÁ EXISTENTE. ALI SERÁ EXIBIDA AS INFORMAÇÕES JA SALVA DO MESMO!

function buscar(){
var ss=SpreadsheetApp.getActive();
var nome= ss.getRange('C3').getValue();
var sslista = ss.getSheetByName('Colaboradores 👨‍🏭');

if (nome == ""){
  SpreadsheetApp.getUi().alert("Funcionario Não Localizado, Favor, selecione um usuario.");
  return false
}

 var ult_func = sslista.getLastRow();
 var func_dados = sslista.getRange(2,1,ult_func-1,7).getValues();

 var dados =[];
 for(let i=0; i< func_dados.length;i++){
  if(func_dados[i][0] == nome){

    for(let y=0; y<7; y++){

       dados.push(func_dados[i][y]);

     }
     break
   }
 }
 //SpreadsheetApp.getUi().alert(dados);

 var ss = ss.getSheetByName('Cadastro 📘');


 ss.getRange('C6').setValue(dados[0]);
 ss.getRange('C9').setValue(dados[1]);
 ss.getRange('C12').setValue(dados[2]);
 ss.getRange('C15').setValue(dados[3]);
 ss.getRange('C18').setValue(dados[4]);
 ss.getRange('C21').setValue(dados[5]);
 ss.getRange('G3').setValue(dados[6]);

}


//FUNÇÃO PARA ALTERAR DADOS JÁ REGISTRADO, SERÁ USADO EM CONJUNTO COM A "MP" PARA CASO UM FUNCIONARIO SEJÁ PROMOVIDO OU 
//TROQUE DE FUNÇÃO. 

function alterar_dados(){

var ssformulario = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Cadastro 📘');
ssformulario.getRange('C6:C18').clear;

}






function excluir(){
var ss = SpreadsheetApp.getActiveSheet();
var funcionario = ss.getRange('C6').getValue();

var sslista=SpreadsheetApp.getActive().getSheetByName("Colaboradores 👨‍🏭");
var ult_func=sslista.getLastRow();
//var func_dados=sslista.getRange(2,1,ult_func-1,1).getValues;

var ui = SpreadsheetApp.getUi();
var resposta =ui.alert('ATENÇÃO!','Deseja excluir os dados de ' + funcionario + ' ? ',ui.ButtonSet.YES_NO); 

if (resposta == "NO" || resposta == "CANCEL"){
  //CASO A EXCLUSÃO SEJA NEGADA

ui.alert("Exclusão cancelada")
return
}else if(resposta == "YES"){

//CASO A EXCLUSÃO SEJA ACEITA

var dados = sslista.getRange(2,1,ult_func).getValues();

for(var linha = 0;linha < dados.length; linha ++){
  if(dados[linha][0]== funcionario){

    var data_aniv = ss.getRange('C12').getValue();
    var id = ss.getRange('G3').getValue();
    var eventDeleted = 0
    //event =calendario.getEventsForDay(data_aniv);
    event = calendario.getEventById(id);
 
    Logger.log(event);
    //ui.alert(event);

    //  event.forEach(event =>{
    //    try{event.deleteEvent()}
    //    finally{eventDeleted ++} 
    //  })


      event.deleteEvent()
      
    

       ss.getRange('C6').clearContent();
       ss.getRange('C9').clearContent();
       ss.getRange('C12').clearContent();
       ss.getRange('C15').clearContent();
       ss.getRange('C18').clearContent();
       ss.getRange('C21').clearContent();
       ss.getRange('G3').clearContent();

      var linha = linha + 2;
      sslista.deleteRow(linha);
      ui.alert("Dados deletados com sucesso!")
      return
    
      }
    }
  }
}




function r(){
var ss = SpreadsheetApp.getActiveSheet();
var funcionario = ss.getRange('C6').getValue();

var sslista=SpreadsheetApp.getActive().getSheetByName("Colaboradores 👨‍🏭");
var ult_func=sslista.getLastRow();
//var func_dados=sslista.getRange(2,1,ult_func-1,1).getValues;

var ui = SpreadsheetApp.getUi();
var resposta =ui.alert('ATENÇÃO!','Deseja excluir os dados de ' + funcionario + ' ? ',ui.ButtonSet.YES_NO); 

if (resposta == "NO" || resposta == "CANCEL"){
  //CASO A EXCLUSÃO SEJA NEGADA

ui.alert("Exclusão cancelada")
return
}else if(resposta == "YES"){

//CASO A EXCLUSÃO SEJA ACEITA

  for (let i = 0; i< funcdados.length ; i++){
     if(funcdados[i][1] == funcionario){
       for (let y=0;y<5;y++){
          sslista.getRange(i+2,y+1).setValue(dadosnovos[y]);
        }
         sslista.getRange(i+2,7);
         ui.alert('Dados alterados com sucesso!');
         return
       // console.log(sslista)
      
    

      ss.getRange('C6').clearContent();
      ss.getRange('C9').clearContent();
      ss.getRange('C12').clearContent();
      ss.getRange('C15').clearContent();
      ss.getRange('C18').clearContent();
      ss.getRange('C21').clearContent();
      ss.getRange('G3').clearContent();

      var linha = linha + 2;
      sslista.deleteRow(linha);
      ui.alert("Dados deletados com sucesso!")
      return
    
      }
    }
  }
}


function alterar_dados(){

var ss = SpreadsheetApp.getActiveSheet();
var funcionario = ss.getRange('C3').getValue();
var dadosnovos = [];
var data = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy")
for(let i=0;i<21;i+=3){
  let item = ss.getRange (i+3,3).getValue();
  dadosnovos.push(item)
}

dadosnovos.push(data);

var ui = SpreadsheetApp.getUi();
var resposta = ui.alert('ATENÇÃO','Deseja alterar os dados de '+funcionario+ '?' ,ui.ButtonSet.YES_NO);

var sslista= SpreadsheetApp.getActive().getSheetByName('Colaboradores 👨‍🏭');
var ult_func = sslista.getLastRow();
var funcdados = sslista.getRange(2,1,ult_func-1,8).getValues();

//ui.alert(ult_func);

  if (resposta == "NO" || resposta == "CANCEL"){
    ui.alert("operação cancelada.")
      
    }
  else{
        for (let i = 0; i< funcdados.length ; i++){
     if(funcdados[i][1] == funcionario){
       for (let y=0;y<5;y++){
          sslista.getRange(i+2,y+1).setValues(dadosnovos[y]);
        }
         sslista.getRange(i+2,7);
         ui.alert('Dados alterados com sucesso!');
         return
       // console.log(sslista)
    }

   }

  }

 }


