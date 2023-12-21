# Projeto-Google-Apps-Scripts
Projeto feito junto com o curso Introduction To Google Apps Script
/**
 * 
 * Introdução ao projeto Apps Script
 * Projeto: codando questionario do Google Forms
 * 
 * Introduction to the Apps Script project
 * Project: coding Google Forms questionnaire
 * 
 * Name: Vitor da Silva Avellar
 */

// conectar com o Google Form
//ID do form
//https://docs.google.com/forms/d/1bSSp7SsecgACWlZKJhWnQDL3u_YRakROQBqyXjMTyyM/edit

const FORM_ID = '1bSSp7SsecgACWlZKJhWnQDL3u_YRakROQBqyXjMTyyM';

//Adicionar um menu customizado para a planilha

function onOpen() {

  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Questionaire Menu')
  .addItem('Atualizar Formulario', 'uptadeForm_V2')
  .addItem('Mandar Emails', 'sendEmail_V2')
  .addToUi();

}

//Pegar os IDs dos componentes do formulario
function getFormIDs() {

  const form = FormApp.openById(FORM_ID);
  const formItems = form.getItems(); // array dos itens dos formularios

  //loop sobre o array
  //imprimir o título e o ID dos itens do formulário
  formItems.forEach(item => console.log(item.getTitle() + '' + item.getId()))

}
/*

11:19:06	Informação	Nome736070830
11:19:06	Informação	Endereço de E-mail546644800
11:19:06	Informação	Você tem alguma experiencia com programação?1758501086
11:19:06	Informação	Quais linguagens de programação você conhece?1722282366
*/


//Atualizando o fomulario pelo Google Sheets
//versão 1
function uptadeForm_V2() {

  // pegar a planilha e as variaveis do formulario
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const form = FormApp.openById(FORM_ID);

  // pegar lista de linguagens da planilha de setup
  const setupSheet = ss.getSheetByName('setup');
  const setupSheetValues = setupSheet.getRange(2, 1, setupSheet.getLastRow() - 1, 1).getValues().flat();
  // console.log(setupSheetValues);
  //[ 'None', 'Apps Script' ]

  // pegar lista de linguagens enviada via formulario
  const responseSheet = ss.getSheetByName('Respostas ao formulário 1');
  const data = responseSheet.getRange(2, 5, responseSheet.getLastRow() - 1, 1).getValues().flat();
  const submittedFormValues = data.join().split(',');
  //console.log(submittedFormValues);
  //[ 'Nenhuma', 'Apps Script', 'Apps Script', ' Python' ]

  // pegar lista de linguagens da questão de formulario
  const formCheckboxChoices = form.getItemById('1722282366').asCheckboxItem().getChoices();
  const formCheckboxValues = formCheckboxChoices.map(x => x.getValue())
  //console.log(formCheckboxValues);
  //[ 'None', 'Apps Script' ]

  //consolidar lista de linguagens
  const allLangs = [...formCheckboxValues, ...setupSheetValues, ...submittedFormValues];
  //console.log(allLangs);
  /*[ 'None',
  'Apps Script',
  'None',
  'Apps Script',
  'Nenhuma',
  'Apps Script',
  Apps Script',
  ' Python' ] */

  //remover o lider e os espaços das linguagens
  const trimLangs = allLangs.map(item => item.trim());
  //console.log(trimLangs);
  /*[ 'None',
  'Apps Script',
  'None',
  'Apps Script',
  'Nenhuma',
  'Apps Script',
  'Apps Script',
  'Python' ] */

  // organizar lista de linguagens
  trimLangs.sort();
  //console.log(trimLangs);
  /*[ 'Apps Script',
  'Apps Script',
  'Apps Script',
  'Apps Script',
  'Nenhuma',
  'None',
  'None',
  'Python' ]*/

  // evitar duplicaçao de linguangens
  let finalLangList = trimLangs.filter((lang, i) => trimLangs.indexOf(lang) === i);
  //console.log(finalLangList)
  //[ 'Apps Script', 'Nenhuma', 'None', 'Python' ]

  // remover os espaços em branco
  finalLangList = finalLangList.filter(item => item.lenght !== 0);

  //[ 'Apps Script', 'Nenhuma', 'Python' ]

  //mover o 'nenhuma' para o topo da array
  finalLangList = finalLangList.filter(item => item !== 'Nenhuma');
  //[ 'Apps Script', 'Python' ]
  finalLangList.unshift('Nenhuma')
  //console.log(finalLangList)
  //[ 'Nenhuma', 'Apps Script', 'Python' ]


  // transformar em arry duplo para planilha
  const finalDoubleArray = finalLangList.map(lang => [lang]);
  //console.log(finalDoubleArray)
  //[ [ 'Nenhuma' ], [ 'Apps Script' ], [ 'Python' ] ]

  //colar dentro da planilha na aba setup
  setupSheet.getRange("A2:A").clear();
  setupSheet.getRange(2,1,finalLangList.length, 1).setValues(finalDoubleArray);

  // copiar dentro do formulario
  form.getItemById('1722282366').asCheckboxItem().setChoiceValues(finalLangList);








}









//Atualizando o fomulario pelo Google Sheets
//versão 1
function uptadeForm_V1() {
  // pegar a lista de linguagens do google sheets
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const setupSheet = ss.getSheetByName('setup');


  const langValues = setupSheet.getRange(2, 1, setupSheet.getLastRow() - 1, 1).getValues();
  console.log(langValues);

  // Informação	[ [ 'None' ], [ 'Apps Script' ] ]

  const langValsFlat = langValues.map(item => item[0]); // ['nome'] => 'nome'
  console.log(langValsFlat);

  // Informação	[ 'None', 'Apps Script' ] //flatened Array

  // pegar o formulario e as questões
  const form = FormApp.openById(FORM_ID);
  const langChecBoxQuestion = form.getItemById('1722282366').asCheckboxItem();

  // preencha a pergunta do formulário com a lista de idiomas
  // array de Strins
  // ['casts', 'dogs']
  langChecBoxQuestion.setChoiceValues(langValsFlat)


}



//version 2
// automaticamente enviar emails para responder com suas informações

function sendEmail_V2() {

  //pegar a informação do spreadSheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const responseSheet = ss.getSheetByName('Respostas ao formulário 1');
  const data = responseSheet.getDataRange().getValues();
  //console.log(data)

  //remover o cabeçalho 
  data.shift();
  //console.log(data);


  //loop nas linhas
  data.forEach((row, i) => {


    //indentificar os que não tiveram replies
    if (row[5] === '') {

      //pegar os endereços de emails
      const name = row[1]
      const email = row[2];
      const answer = row[3];// sim/nao
      const langs = row[4]; // lista de linguagens


      // escrever os emails
      const subject = 'obrigado por responder o questionario Apps Script '
      let body = '';

      //mudar o body do email caso seja 'sim' ou 'nao'
      if (answer === 'Sim') {
        body = 'Olá ' + name + `, <br><br>
        Obrigado por responder a pesquisa!<br><br>
        Seu feedBack é muito importante.<br><br>
        voce reportou experiencias com as linguagens:<br><br>
       <em>`+ langs + `</em><br><br>
       Obrigado,
       Vitor`;
      }//sem resposta
      else {
        body = 'Olá ' + name + `, <br><br>
        Obrigado por responder a pesquisa!<br><br>
        Seu feedBack é muito importante.<br><br>
        voce reportou que não tem experiencias com linguagens de programação, aqui está algumas sugestões para começar: <br><br>
        <a href="https://courses.benlcollins.com/courses/enrolled/435404">Começe aqui com AppsScripts</a><br><br>
        Obrigado,<br>
       <em>`;

      }
      // console.log(email)
      // console.log(subject)
      // console.log(body)


      //enviar os emails
      GmailApp.sendEmail(email, subject, '', { htmlBody: body });

      //marcar como enviado
      const d = new Date()
      responseSheet.getRange(i + 2, 6).setValue(d);
    }
    else {
      console.log('Sem Email nessa linha')
    }
  })
}




//version 1
// automaticamente enviar emails para responder com suas informações
function sendEmail_V1() {

  //pegar a informação do spreadSheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const responseSheet = ss.getSheetByName('Respostas ao formulário 1');
  const data = responseSheet.getDataRange().getValues();
  //console.log(data)

  //remover o cabeçalho 
  data.shift();
  //console.log(data);


  //loop nas linhas
  data.forEach((row, i) => {


    //indentificar os que não tiveram replies
    if (row[5] === '') {



      //pegar os endereços de emails
      const email = row[2]




      // escrever os emails
      const subject = 'obrigado por responder o questionario Apps Script '
      let body = '';

      //mudar o body do email caso seja 'sim' ou 'nao'
      if (row[3] === 'Sim') {
        body = 'TBC - resposta positiva';
      }//sem resposta
      else {
        body = 'TBC - resposta negativa';

      }



      //enviar os emails
      GmailApp.sendEmail(email, subject, body);

      //marcar como enviado
      const d = new Date()
      responseSheet.getRange(i + 2, 6).setValue(d);
    }
    else {
      console.log('Sem Email nessa linha')
    }
  })
}














