function getParentFolderByName_(folderName) {

  // Gets the Drive Folder of where the current spreadsheet is located.
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const parentFolder = DriveApp.getFileById(ssId).getParents().next();

  // Iterates the subfolders to check if the PDF folder already exists.
  const subFolders = parentFolder.getFolders();
  while (subFolders.hasNext()) {
    let folder = subFolders.next();

    // Returns the existing folder if found.
    if (folder.getName() === folderName) {
      return folder;
    }
  }
  // Creates a new folder if one does not already exist.
  return parentFolder.createFolder(folderName)
    .setDescription(`Created by Eu mesmo: application to store PDF output files`);
}

const plan_atual = SpreadsheetApp.getActiveSpreadsheet()
const aba_admp = plan_atual.getSheetByName('ADMP')
const FORMULARIO = plan_atual.getSheetByName('formRecisao')
const folder = getParentFolderByName_('pdfs')
const ssId = plan_atual.getId()

//PLANILHA MANUTENÇÃO ------------------------------------------------------------------------

const SHEET_MANUTENÇÃO = SpreadsheetApp.openById("1muBwSOryC27MtuRXkhpxSOdnFPz46uIeMdQcGrr8O4w").getSheetByName("Respostas form. resc.")

//PLANILHA DAS REGIONAIS ------------------------------------------------------------------------

const SHEETS_NORTE = SpreadsheetApp.openById("1DfASkX32WMqTINpwVqQB1tI7TIGqIrZoqEOV0SmOvp0").getSheetByName("Monitoramento de Rescisões - NORTE")
const SHEETS_NORDESTE = SpreadsheetApp.openById("1tjXAJFEpHfVPcI5hKXzaMy78N0Qyf_hyDGCGt0eugGw").getSheetByName("Monitoramento de Rescisões - NORDESTE")
const SHEETS_NOROESTE = SpreadsheetApp.openById("14JHOzFLD2TgboahG6UWq_NUQwTBSKRL8ciu9Ui8TNkc").getSheetByName("Monitoramento de Rescisões - NOROESTE")
const SHEETS_OESTE = SpreadsheetApp.openById("1tdwi7eeduxzS5nD1m1QEKAVZMFKDJdVW5vbr1KASU84").getSheetByName("Monitoramento de Rescisões - OESTE")
const SHEETS_CENTRO_SUL = SpreadsheetApp.openById("1gErtvhoWUWRB_NpzMb1OUEQPmFYn5lNmHaNDYaGpMYI").getSheetByName("Monitoramento de Rescisões - CENTRO SUL")
const SHEETS_LESTE = SpreadsheetApp.openById("1xJu7v6PNE1TPd-8yo4PnwGR2ClGmVnyi9AfTYjLTInA").getSheetByName("Monitoramento de Rescisões - LESTE")
const SHEETS_NIVEL_CENTRAL = SpreadsheetApp.openById("12NX5qxba9wKXmFt5Cng7kSSmoRghOUgRF5nQEj-4t4k").getSheetByName("Monitoramento de Rescisões - NÍVEL CENTRAL")
const SHEETS_BARREIRO = SpreadsheetApp.openById("1DagPF0PbCE7Ed1F6A60yV4pfMuig2faqRwY-2Azry8g").getSheetByName("Monitoramento de Rescisões - BARREIRO")
const SHEETS_VENDA_NOVA = SpreadsheetApp.openById("19E00acq7fGLTHAUs9_vVwvgeTZVvxKnISzorYzjbvv8").getSheetByName("Monitoramento de Rescisões - VENDA NOVA")
const SHEETS_PAMPULHA = SpreadsheetApp.openById("1WogRODhxMkn_2KP-bVD0Q3d4_YjfHYWcJZEu0yqtN1Q").getSheetByName("Monitoramento de Rescisões - PAMPULHA")

//----------------------------------------------------------------------------------------------------

//Link preenchido para testes para email Artur
//https://docs.google.com/forms/d/e/1FAIpQLSdj3wmwOP_le-fIQy_4xoTS90bYSBhAqP3rkj_7CH_XX84WjQ/viewform?usp=pp_url&entry.1425630199=arturfmmendes@gmail.com&entry.1308478672=Por+interesse+da+SMSA&entry.1929694843=N%C3%A3o+atende+%C3%A0s+expectativas+do+servi%C3%A7o&entry.1344356617=0134604&entry.1258792198=08022811688&entry.1817285980=111111&entry.1386471171=2022-08-04&entry.322925036=CIENTE
//OK
function onFormSubmit(e) {

  const RespostasForms = e.namedValues
  
  //Pega os dados do ADMP da matricula e confere se o CPF bate
  const DadosADMP = capturaDados(RespostasForms)

  //se o CPF não bater
  if(DadosADMP.erro === 1){
    enviarEmaildeErro(DadosADMP)
    Logger.log("CPF e BM não são correspondentes: Email de erro enviado ao destinatário " + DadosADMP.email)
    return 1
  }

  //se o e-mail do gestor e do profissional forem iguais
  if(DadosADMP.erro === 2){
    enviarEmaildeErroEmail(DadosADMP)
    Logger.log("E-mail não enviado: e-mail do gestor e do profissional são iguais")
    return 1
  }
  
  //atribui os valores de todos os dados necessarios a um objeto
  const DadosProfissional = atribuirValores(DadosADMP, RespostasForms)


  //distribui os dados do objeto pelo termo de rescisao
  preencherPlanilha(DadosProfissional)

  //gera o pdf
  const PDF = gerarPDF(DadosProfissional.nome)

  //envia email com o objeto a seguir
  enviarEmail(
    {
      confirmaBV: DadosProfissional.confirmaBV,
      reciboBV: DadosProfissional.reciboBV,
      emailGestor: DadosProfissional.emailGestor,
      emailProfissional: DadosProfissional.emailProfissional,
      nome: DadosProfissional.nome.toString(),
      matricula: DadosProfissional.matricula.toString(),
      ultDia: DadosProfissional.ultimoDia.toString(),
      interesse: DadosProfissional.iniciativaRecisao.toString(),
      cienciaProfissional: RespostasForms["Declaro que já comuniquei ao profissional quanto a rescisão de seu contrato."].toString(),
      pdf: PDF
    }
  )

  folder.getFilesByName(PDF.getName()).next().setTrashed(true)

  envioManutencao(DadosProfissional, RespostasForms)
  
  envioUnidades(DadosProfissional, RespostasForms)

  
}

//OK
function capturaDados(FormRespostas){
  
  Logger.log("Resposta recebida")
  Logger.log(JSON.stringify(FormRespostas))

  const matriculaEnviada = FormRespostas['MATRÍCULA DO PROFISSIONAL']
  const cpfEnviado = FormRespostas['CPF DO PROFISSIONAL'].toString()
  const emailGestor = FormRespostas['E-MAIL DO GESTOR'].toString()
  var emailProfissional = FormRespostas['E-MAIL DO PROFISSIONAL'][0].toString()

  if(emailProfissional === ""){
    emailProfissional = FormRespostas['E-MAIL DO PROFISSIONAL'][1].toString()
  }
  
  //linha dos dados do profissional no ADMP
  const matrizADMP = aba_admp.getDataRange().getValues()
  const profissionalADMP = matrizADMP.filter(
    data => {
    if(data[0] == matriculaEnviada){
      return data
    }
  })[0]

  if (profissionalADMP[2] != cpfEnviado){
    return { erro: 1, emailGestor, matrEnviada: matriculaEnviada, cpfEnv: cpfEnviado}
  }

  //passando os 2 para LowerCase para não aceitar o mesmo email maiúsculo e minusculo
  if(emailGestor.toLowerCase() === emailProfissional.toLowerCase()){
    return { erro: 2, emailGestor, matrEnviada: matriculaEnviada, cpfEnv: cpfEnviado}
  }
  
  return profissionalADMP
  
}

//OK
function preencherPlanilha(DadosProf){

  Logger.log("Dados recebidos para preenchiento do termo")
  //Logger.log(JSON.stringify(DadosProf))
  
  //pega o nome da celula
  FORMULARIO.getNamedRanges().forEach(
    
    namedRange => {
    const nomeArr = namedRange.getName().split("!")
    let nome
    if(nomeArr.length === 2){
      nome = nomeArr[1]
    }else{
      nome = nomeArr[0]
    }

    switch(nome){
      case 'form_nome':
        namedRange.getRange().setValue(DadosProf.nome)
        break;
      case 'form_cpf':
        namedRange.getRange().setValue(DadosProf.cpf)
        break;
      case 'form_matricula':
        namedRange.getRange().setValue(DadosProf.matricula)
        break;
      case 'form_categoria':
        if (DadosProf.especialidade === "" || DadosProf.especialidade === "MEDICO") {
        
          namedRange.getRange().setValue(DadosProf.categoria)
            
        } else {
        
          namedRange.getRange().setValue(DadosProf.categoria + " / " + DadosProf.especialidade)
        }
        break;
      case 'form_nRegistro':
        if(DadosProf.numRegistro === ""){
          
          namedRange.getRange().setValue("NÃO SE APLICA")
        }
        else{
          namedRange.getRange().setValue(DadosProf.numRegistro)
        }
        break;
      case 'form_conselho':
        if(DadosProf.conselho === ""){
          
          namedRange.getRange().setValue("NÃO SE APLICA")
        }
        else{
          namedRange.getRange().setValue(DadosProf.conselho)
        }

        break;
      case 'form_ufConselho':
        if(DadosProf.ufConselho === "" || DadosProf.ufConselho === "-"){
          
          namedRange.getRange().setValue("NÃO SE APLICA")
        }
        else{
          namedRange.getRange().setValue(DadosProf.ufConselho)
        }

        break;
      case 'form_inicioContrato':
        namedRange.getRange().setValue(DadosProf.dataAdmissao)
        break;
      case 'form_regional1':
        namedRange.getRange().setValue(DadosProf.regional1)
        break;
      case 'form_lotacao1':
        namedRange.getRange().setValue(DadosProf.lotacao1)
        break;
      case 'form_ultDia':
        namedRange.getRange().setValue(DadosProf.ultimoDia)
        break;    
      case 'form_bv':
        if(DadosProf.reciboBV === ""){
          namedRange.getRange().setValue("ENVIO MANUAL")
        }else{
          namedRange.getRange().setValue(DadosProf.reciboBV)
        }
        
        namedRange.getRange().setNumberFormat('000000')
        break;
      case 'form_intProf':
        if(DadosProf.iniciativaRecisao == 'Por interesse do profissional'){
          namedRange.getRange().setValue(true)
        }else{
          namedRange.getRange().setValue(false)
        }
        break;
      case 'form_intSMSA':
        if(DadosProf.iniciativaRecisao == 'Por interesse da SMSA'){
          namedRange.getRange().setValue(true)
        }else{
          namedRange.getRange().setValue(false)
        }
        break;
    
     
      default:
        namedRange.getRange().setValue("")
    }
  })
  Logger.log("Termo preenchido")
}

//OK
function gerarPDF(nomeProfissional){

  Logger.log('Iniciando impressão em PDF')

  console.log(FORMULARIO.getRange('b16').getValue() + " Atraso necessário")
  
  const fr = 0, fc = 0, lr = 67, lc = 12
  const url = "https://docs.google.com/spreadsheets/d/" + ssId + "/export" +
    "?format=pdf&" +
    "size=7&" +
    "fzr=true&" +
    "portrait=true&" +
    "fitw=true&" +
    "gridlines=false&" +
    "printtitle=false&" +
    "top_margin=0.2&" +
    "bottom_margin=0.2&" +
    "left_margin=0.2&" +
    "right_margin=0.2&" +
    "sheetnames=false&" +
    "pagenum=UNDEFINED&" +
    "attachment=true&" +
    "gid=" + FORMULARIO.getSheetId() + '&' +
    "r1=" + fr + "&c1=" + fc + "&r2=" + lr + "&c2=" + lc;


  const params = { method: "GET", headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() } };
  const blob = UrlFetchApp.fetch(url, params).getBlob().setName(nomeProfissional + "---" + hoje() + '.pdf');

  const pdf = folder.createFile(blob)
  Logger.log('PDF criado')
  return pdf
}

//OK
function excluirPDF(pdf) {
  Logger.log('excluindo PDF')
  folder.getFilesByName(pdf.getName()).next().setTrashed(true)
  Logger.log("PDF excluído.")
}

var image = DriveApp.getFileById("1kOoCMcm57rZMwa71OgyeOalZpSddirrs").getAs("image/png");
var emailImages = {"logo": image};


function enviarEmail(detalhesEmail){


  const CORPO_TEXTO_PROFISSIONAL = corpoTextoProfissional(detalhesEmail.nome, detalhesEmail.matricula, detalhesEmail.ultDia,detalhesEmail.confirmaBV, detalhesEmail.interesse, detalhesEmail.cienciaProfissional)
  const CORPO_TEXTO_GESTOR = corpoTextoGestor(detalhesEmail.nome, detalhesEmail.reciboBV, detalhesEmail.confirmaBV)
  
  // email para o gestor
  GmailApp.sendEmail(
    //emails
    detalhesEmail.emailGestor,

    //título
    'DIEP | GGCAT | Rescisão CADM de ' + primeirosNomes(detalhesEmail.nome) + " em " + hoje() + " | Mat.: " + detalhesEmail.matricula,

    //corpo do email 
    
    'corpoEmail',
    
    //anexos do email
    {
      attachments: [detalhesEmail.pdf.getAs(MimeType.PDF)],
      name: "Rescisão de CADM",
      htmlBody: CORPO_TEXTO_GESTOR,
      inlineImages: emailImages
    }
  )
  Logger.log("Email para o gestor enviado com sucesso.");
  
  // email para o profissional
  if(detalhesEmail.interesse === "Por interesse do profissional" || detalhesEmail.confirmaBV === "Não"){
    GmailApp.sendEmail(
      //emails
      detalhesEmail.emailProfissional.toString(),

      //título
      'DIEP | GGCAT | Rescisão CADM de ' + primeirosNomes(detalhesEmail.nome) + " em " + hoje() + " | Mat.: " + detalhesEmail.matricula,

      //corpo do email 
      
      'corpoEmail',
      
      //anexos do email
      {
        name: "Rescisão de CADM",
        htmlBody: CORPO_TEXTO_PROFISSIONAL,
        inlineImages: emailImages
      }
    )
    Logger.log("Email para o profissional enviado com sucesso.")
  }
}

function enviarEmaildeErro(CPFeMatriculaErrados){
  
  //protegendo o cpf pegando os 3 digitos do meio
  var cpf3 = CPFeMatriculaErrados.cpfEnv.toString()
  cpf3 = cpf3[3] + cpf3[4] + cpf3[5]

  //email html
  var corpoHtml = '<html> <head> <meta charset="UTF-8"> </head><style> body{ font-family:Verdana, Geneva, Tahoma, sans-serif; } </style> <body><p>' +
  saudacao() + '.<br><br>' +

  "O Termo de rescisão não pôde ser gerado pois a Mat.: " + CPFeMatriculaErrados.matrEnviada + " e o CPF: ***." + cpf3 + ".***-** não correspondem ao mesmo contrato. Favor realizar outra submissão do formulário com dados corretos.<br><br>" +
  'Caso esta mensagem não faça sentido para você, gentileza desconsiderá-la.<br><br>' + 'Atenciosamente,<br><br></p>' + assinatura

  GmailApp.sendEmail(
      //destinatários
      CPFeMatriculaErrados.emailGestor,
      //título
      'DIEP | GGCAT | Rescisão CADM | ERRO NA EMISSÃO DO TERMO DE RESCISÃO',
      //corpo do email
      'corpoHtml',

      {
        name: "Rescisão de CADM",
        htmlBody: corpoHtml,
        inlineImages: emailImages
      }
    )
}

function enviarEmaildeErroEmail(CPFeMatriculaErrados){
  
  //protegendo o cpf pegando os 3 digitos do meio
  var cpf3 = CPFeMatriculaErrados.cpfEnv.toString()
  cpf3 = cpf3[3] + cpf3[4] + cpf3[5]

  //email html
  var corpoHtml = '<html> <head> <meta charset="UTF-8"> </head><style> body{ font-family:Verdana, Geneva, Tahoma, sans-serif; } </style> <body><p>' +
  saudacao() + '.<br><br>' +

  "O Termo de rescisão não pôde ser gerado pois os e-mails informados como o do gestor e o do profissional são iguais. Favor realizar outra submissão do formulário com dados corretos.<br><br>" +
  'Caso esta mensagem não faça sentido para você, gentileza desconsiderá-la.<br><br>' + 'Atenciosamente,<br><br></p>' + assinatura

  GmailApp.sendEmail(
      //destinatários
      CPFeMatriculaErrados.emailGestor,
      //título
      'DIEP | GGCAT | Rescisão CADM | ERRO NA EMISSÃO DO TERMO DE RESCISÃO',
      //corpo do email
      'corpoHtml',

      {
        name: "Rescisão de CADM",
        htmlBody: corpoHtml,
        inlineImages: emailImages
      }
    )
}



function envioManutencao(Profissional, Respostas){


  const valores = [
    [Respostas["Carimbo de data/hora"], Respostas["E-MAIL DO GESTOR"], Respostas["INICIATIVA DA RESCISÃO"], Respostas["MOTIVOS DA RESCISÃO"], Profissional.emailProfissional, Profissional.matricula, Profissional.cpf, Profissional.ultimoDia, Profissional.reciboBV, Profissional.matricula, Profissional.nome, Profissional.lotacao1, Profissional.regional1, Profissional.categoria, Profissional.dataAdmissao]
  ]


  //ENVIO PRA MANUTANCAO
  
  //passar pela coluna 11 
  const colunaK = SHEET_MANUTENÇÃO.getRange("K:K").getValues()
  var ultimaLinha = "" 

  for(let i = 0; i < colunaK.length; i++){

    if(colunaK[i][0] == ""){

      ultimaLinha = i

      break;
    }
  }

  SHEET_MANUTENÇÃO.getRange(ultimaLinha + 1, 1, 1, 15).setValues(valores)

  Logger.log("Enviado para planilha de rescisões: " + valores)
}




function envioUnidades(Profissional, Respostas){
  
  const valores = [
    [Profissional.matricula, Profissional.cpf, Profissional.nome, Profissional.lotacao1, Profissional.regional1, Profissional.categoria, Profissional.dataAdmissao, Profissional.ultimoDia, Profissional.reciboBV, Respostas["Carimbo de data/hora"], Respostas["E-MAIL DO GESTOR"], Profissional.emailProfissional, Respostas["INICIATIVA DA RESCISÃO"], Respostas["MOTIVOS DA RESCISÃO"]]
  ]

  var colunaB = ""
  var ultimaLinha = "" 

  switch(Profissional.regional1){
    
    case "BARREIRO":
      
      colunaB = SHEETS_BARREIRO.getRange("B3:B").getValues()
      ultimaLinha = descobrirUltimaLinha(colunaB)
      SHEETS_BARREIRO.getRange(ultimaLinha + 1, 2, 1, 14).setValues(valores)
      Logger.log("Enviado à regional BARREIRO")
      break;

    case "CENTRO SUL":
      
      colunaB = SHEETS_CENTRO_SUL.getRange("B3:B").getValues()
      ultimaLinha = descobrirUltimaLinha(colunaB)
      SHEETS_CENTRO_SUL.getRange(ultimaLinha + 1, 2, 1, 14).setValues(valores)
      Logger.log("Enviado à regional CENTRO SUL")
      break;

    case "LESTE":
      
      colunaB = SHEETS_LESTE.getRange("B3:B").getValues()
      ultimaLinha = descobrirUltimaLinha(colunaB)
      SHEETS_LESTE.getRange(ultimaLinha + 1, 2, 1, 14).setValues(valores)
      Logger.log("Enviado à regional LESTE")
      break;

    case "NORTE":
      
      colunaB = SHEETS_NORTE.getRange("B3:B").getValues()
      ultimaLinha = descobrirUltimaLinha(colunaB)
      SHEETS_NORTE.getRange(ultimaLinha + 1, 2, 1, 14).setValues(valores)
      Logger.log("Enviado à regional NORTE")
      break;

    case "NÍVEL CENTRAL":
      
      colunaB = SHEETS_NIVEL_CENTRAL.getRange("B3:B").getValues()
      ultimaLinha = descobrirUltimaLinha(colunaB)
      SHEETS_NIVEL_CENTRAL.getRange(ultimaLinha + 1, 2, 1, 14).setValues(valores)
      Logger.log("Enviado à regional CENTRO SUL")
      break;

    case "NORDESTE":
      
      colunaB = SHEETS_NORDESTE.getRange("B3:B").getValues()
      ultimaLinha = descobrirUltimaLinha(colunaB)
      SHEETS_NORDESTE.getRange(ultimaLinha + 1, 2, 1, 14).setValues(valores)
      Logger.log("Enviado à regional NORDESTE")
      break;
    
    case "NOROESTE":
      
      colunaB = SHEETS_NOROESTE.getRange("B3:B").getValues()
      ultimaLinha = descobrirUltimaLinha(colunaB)
      SHEETS_NOROESTE.getRange(ultimaLinha + 1, 2, 1, 14).setValues(valores)
      Logger.log("Enviado à regional NOROESTE")
      break;
    
    case "VENDA NOVA":
      
      colunaB = SHEETS_VENDA_NOVA.getRange("B3:B").getValues()
      ultimaLinha = descobrirUltimaLinha(colunaB)
      SHEETS_VENDA_NOVA.getRange(ultimaLinha + 1, 2, 1, 14).setValues(valores)
      Logger.log("Enviado à regional VENDA NOVA")
      break;

    case "PAMPULHA":
      
      colunaB = SHEETS_PAMPULHA.getRange("B3:B").getValues()
      ultimaLinha = descobrirUltimaLinha(colunaB)
      SHEETS_PAMPULHA.getRange(ultimaLinha + 1, 2, 1, 14).setValues(valores)
      Logger.log("Enviado à regional PAMPULHA")
      break;

    default:
      Logger.log("Nenhuma regional endereçada.")
      break;
  }

  Logger.log("Valores enviados: " + valores)
}
  
function descobrirUltimaLinha(arrayValores){

  var linha = ""

  for(let i = 0; i < arrayValores.length; i++){

    if(arrayValores[i][0] == ""){

      linha = i + 2
      break;
    }
 
  }
  return linha
}  
  