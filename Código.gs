//VARIÁVEIS GLOBAIS ------------------------------------------------------------------------

const plan_atual = SpreadsheetApp.getActiveSpreadsheet()
const aba_admp = plan_atual.getSheetByName('ADMP')
var FORMULARIO
const folder = getParentFolderByName_('pdfs')
const respostasUnidade = getParentFolderByName_('Respostas por unidade')
const ssId = plan_atual.getId()
var image = DriveApp.getFileById("1kOoCMcm57rZMwa71OgyeOalZpSddirrs").getAs("image/png");
var emailImages = {"logo": image};

//PLANILHA MANUTENÇÃO ------------------------------------------------------------------------

const SHEET_MANUTENÇÃO = SpreadsheetApp.openById("1muBwSOryC27MtuRXkhpxSOdnFPz46uIeMdQcGrr8O4w").getSheetByName("Respostas form. resc.")

//----------------------------------------------------------------------------------------------------

//Link preenchido para testes para email Artur
//https://docs.google.com/forms/d/e/1FAIpQLSdj3wmwOP_le-fIQy_4xoTS90bYSBhAqP3rkj_7CH_XX84WjQ/viewform?usp=pp_url&entry.1425630199=arturfmmendes@gmail.com&entry.1308478672=Por+interesse+da+SMSA&entry.1929694843=N%C3%A3o+atende+%C3%A0s+expectativas+do+servi%C3%A7o&entry.1344356617=0134604&entry.1258792198=08022811688&entry.1817285980=111111&entry.1386471171=2022-08-04&entry.322925036=CIENTE
//OK

function onFormSubmit(e) {

  const RespostasForms = e.namedValues
  console.log(JSON.stringify(RespostasForms))
  
  //Pega os dados do ADMP da matricula e confere se o CPF bate
  const DadosADMP = capturaDados(RespostasForms)

  console.log("DadosADMP")
  console.log(JSON.stringify(DadosADMP))

  //se o CPF não bater
  if(DadosADMP.erro === 1){
    const msg = ""
    const codErr = "erro1"
    enviarEmaildeErro(DadosADMP, DadosADMP.emailGestor,msg, codErr)
    Logger.log("CPF e BM não são correspondentes: Email de erro enviado ao destinatário " + DadosADMP.emailGestor)
    return 1
  }

  //se o e-mail do gestor e do profissional forem iguais
  if(DadosADMP.erro === 2){
    const msg = ""
    const codErr = "erro2"
    enviarEmaildeErro(DadosADMP, DadosADMP.emailGestor,msg, codErr)
    Logger.log("E-mail não enviado: e-mail do gestor e do profissional são iguais")
    return 2
  }

  //atribui os valores de todos os dados necessarios a um objeto
  const DadosProfissional = atribuirValores(DadosADMP, RespostasForms)

  const dataLei = new Date(2022,6,1)

  //se a contratação do profissional ocorreu a partir de 01/07/2022 e mas foi marcado que iria cumprir aviso prévio
  if(DadosProfissional.dataAdmissao >= dataLei && DadosProfissional.avisoPrevio === "Sim."){
    const msg = "a contratação do profissional ocorreu a partir de 01/07/2022, portanto ele não tem direito à aviso prévio, porém, no preenchimento do formulário foi indicado o cumprimento do aviso. Gentileza submeter novo formulário com os dados corretos." 
    const codErr = "erro3"
    DadosADMP.emailGestor = RespostasForms['E-MAIL DO GESTOR'].toString()
    enviarEmaildeErro(DadosADMP, DadosProfissional.emailGestor,msg, codErr)
    Logger.log("não tem direito a aviso prévio")
    return codErr
  }

  if(DadosProfissional.dataAdmissao >= dataLei ){
    FORMULARIO = plan_atual.getSheetByName('formRecisao')
  }else{
    FORMULARIO = plan_atual.getSheetByName('formRecisao-Aviso')
  }

  //distribui os dados do objeto pelo termo de rescisao
  const ultDia = preencherPlanilha(DadosProfissional,FORMULARIO)

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
      ultDia, //vindo do preenchimento do formulário na planilha
      interesse: DadosProfissional.iniciativaRecisao.toString(),
      cienciaProfissional: RespostasForms["Declaro que já comuniquei ao profissional quanto a rescisão de seu contrato."].toString(),
      pdf: PDF
    }
  )

  folder.getFilesByName(PDF.getName()).next().setTrashed(true)

  envioManutencao(DadosProfissional,ultDia, RespostasForms)
  
  envioUnidades(DadosProfissional, ultDia, RespostasForms)

}

function capturaDados(FormRespostas){
  
  Logger.log("Resposta recebida")
  Logger.log(JSON.stringify(FormRespostas))

  const matriculaEnviada = FormRespostas['MATRÍCULA DO PROFISSIONAL']
  const cpfEnviado = FormRespostas['CPF DO PROFISSIONAL'].toString()
  const emailGestor = FormRespostas['E-MAIL DO GESTOR'][0].toString() || FormRespostas['E-MAIL DO GESTOR'][1].toString()
  const emailProfissional = FormRespostas['E-MAIL DO PROFISSIONAL'][0].toString() || FormRespostas['E-MAIL DO PROFISSIONAL'][1].toString()
  
  //linha dos dados do profissional no ADMP
  const matrizADMP = aba_admp.getDataRange().getValues()
  const profissionalADMP = matrizADMP.filter(
    data => {
    if(data[3] == matriculaEnviada){
      return data
    }
  })[0]

  if (profissionalADMP[1] != cpfEnviado){
    return { erro: 1, emailGestor, matriculaEnviada, cpfEnviado}
  }

  //passando os 2 para LowerCase para não aceitar o mesmo email maiúsculo e minusculo
  if(emailGestor.toLowerCase() === emailProfissional.toLowerCase()){
    return { erro: 2, emailGestor, matriculaEnviada, cpfEnviado}
  }
  
  return profissionalADMP
  
}

function preencherPlanilha(DadosProf,FORMULARIO){

  const aba = FORMULARIO.getName()

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
      case 'form_diaAviso':
        if(DadosProf.inicioAviso !== ""){
          namedRange.getRange().setValue(DadosProf.inicioAviso)
        }else{
          namedRange.getRange().setValue(DadosProf.comunicacaoAviso)
        }
        break;
      case `form_nome`:
        namedRange.getRange().setValue(DadosProf.nome)
        break;
      case `form_cpf`:
        namedRange.getRange().setValue(DadosProf.cpf)
        break;
      case `form_matricula`:
        namedRange.getRange().setValue(DadosProf.matricula)
        break;
      case `form_categoria`:
        if (DadosProf.especialidade === "" || DadosProf.especialidade === "MEDICO") {
        
          namedRange.getRange().setValue(DadosProf.categoria)
            
        } else {
        
          namedRange.getRange().setValue(DadosProf.categoria + " / " + DadosProf.especialidade)
        }
        break;
      case `form_nRegistro`:
        if(DadosProf.numRegistro === ""){
          
          namedRange.getRange().setValue("NÃO SE APLICA")
        }
        else{
          namedRange.getRange().setValue(DadosProf.numRegistro)
        }
        break;
      case `form_conselho`:
        if(DadosProf.conselho === ""){
          
          namedRange.getRange().setValue("NÃO SE APLICA")
        }
        else{
          namedRange.getRange().setValue(DadosProf.conselho)
        }

        break;
      case `form_ufConselho`:
        if(DadosProf.ufConselho === "" || DadosProf.ufConselho === "-"){
          
          namedRange.getRange().setValue("NÃO SE APLICA")
        }
        else{
          namedRange.getRange().setValue(DadosProf.ufConselho)
        }

        break;
      case `form_inicioContrato`:
        namedRange.getRange().setValue(DadosProf.dataAdmissao)
        break;
      case `form_regional1`:
        namedRange.getRange().setValue(DadosProf.regional1)
        break;
      case `form_lotacao1`:
        namedRange.getRange().setValue(DadosProf.lotacao1)
        break;  
      case `form_bv`:
        if(DadosProf.reciboBV === ""){
          namedRange.getRange().setValue("ENVIO MANUAL")
        }else{
          namedRange.getRange().setValue(DadosProf.reciboBV)
        }
        
        namedRange.getRange().setNumberFormat('000000')
        break;
      case `form_intProf`:
        if(DadosProf.iniciativaRecisao === 'Por interesse do profissional'){
          namedRange.getRange().setValue(true)
        }else{
          namedRange.getRange().setValue(false)
        }
        break;
      case `form_intSMSA`:
        if(DadosProf.iniciativaRecisao === 'Por interesse da SMSA'){
          namedRange.getRange().setValue(true)
        }else{
          namedRange.getRange().setValue(false)
        }
        break;
      case `form_avisoSIM`:
        if(DadosProf.avisoPrevio === "Sim."){
          namedRange.getRange().setValue(true)
        }
        else{
          namedRange.getRange().setValue(false)
        }
        break;
      case `form_avisoNAO`:
        if(DadosProf.avisoPrevio === "Não."){
          namedRange.getRange().setValue(true)
        }
        else{
          namedRange.getRange().setValue(false)
        }
        break;
      default:
        namedRange.getRange().setValue("")
    }
  })
  
  var celUltDia

  if(aba === "formRecisao"){
    celUltDia = "D31"
    FORMULARIO.getRange(celUltDia).setValue(DadosProf.ultimoDia[0])
  }else{
    celUltDia = "D33"
    const fimAviso = FORMULARIO.getRange("L36").getValue()
    var ultDia = DadosProf.ultimoDia[1] || fimAviso
    FORMULARIO.getRange(celUltDia).setValue(ultDia)
  }

  Logger.log("Formulário preenchido");

  return FORMULARIO.getRange(celUltDia).getValue()

}

function gerarPDF(nomeProfissional){

  Logger.log('Iniciando impressão em PDF')

  console.log(FORMULARIO.getRange('b16').getValue() + " Atraso necessário")
  
  const fr = 0, fc = 0, lr = 67, lc = 13
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

function enviarEmaildeErro(DadosADMP, emailGestor, msg,codErr){

  //email html
  var corpoHtml

  if(codErr === "erro1"){
    //protegendo o cpf pegando os 3 digitos do meio
    var cpf3 = DadosADMP.cpfEnviado.toString()
    cpf3 = cpf3[3] + cpf3[4] + cpf3[5]

    corpoHtml = '<html> <head> <meta charset="UTF-8"> </head><style> body{ font-family:Verdana, Geneva, Tahoma, sans-serif; } </style> <body><p>' +
    saudacao() + '.<br><br>' +

    "O Termo de rescisão não pôde ser gerado pois a Mat.: " + DadosADMP.matriculaEnviada + " e o CPF: ***." + cpf3 + ".***-** não correspondem ao mesmo contrato. Favor realizar outra submissão do formulário com dados corretos.<br><br>" +
    'Caso esta mensagem não faça sentido para você, gentileza desconsiderá-la.<br><br>' + 'Atenciosamente,<br><br></p>' + assinatura
  }

  if(codErr === "erro2"){

    //email html
    var corpoHtml = '<html> <head> <meta charset="UTF-8"> </head><style> body{ font-family:Verdana, Geneva, Tahoma, sans-serif; } </style> <body><p>' +
    saudacao() + '.<br><br>' +

    "O Termo de rescisão não pôde ser gerado pois os e-mails informados como o do gestor e o do profissional são iguais. Favor realizar outra submissão do formulário com dados corretos.<br><br>" +
    'Caso esta mensagem não faça sentido para você, gentileza desconsiderá-la.<br><br>' + 'Atenciosamente,<br><br></p>' + assinatura
  }

  if(codErr === "erro3"){
    //protegendo o cpf pegando os 3 digitos do meio
    cpf3 = DadosADMP[1].toString()
    cpf3 = cpf3[3] + cpf3[4] + cpf3[5]

    const matriculaEnviada = DadosADMP[3].toString()

    corpoHtml = '<html> <head> <meta charset="UTF-8"> </head><style> body{ font-family:Verdana, Geneva, Tahoma, sans-serif; } </style> <body><p>' +
    saudacao() + '.<br><br>' +

    "O Termo de rescisão não pôde ser gerado para o profissional de Mat.: " + matriculaEnviada + " e CPF: ***." + cpf3 + ".***-**, pois " + msg +
    '<br><br>Caso esta mensagem não faça sentido para você, gentileza desconsiderá-la.<br><br>' + 'Atenciosamente,<br><br></p>' + assinatura
  }

  GmailApp.sendEmail(
      //destinatários
      emailGestor,
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

/*function enviarEmaildeErroEmail(CPFeMatriculaErrados){
  
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
}*/

function envioManutencao(DadosProfissional,ultDia, Respostas){
  let ultimoDia = DadosProfissional.ultimoDia[0] || DadosProfissional.ultimoDia[1] || ultDia

  const valores = [
    [Respostas["Carimbo de data/hora"], DadosProfissional.emailGestor, Respostas["INICIATIVA DA RESCISÃO"], Respostas["MOTIVOS DA RESCISÃO"], DadosProfissional.emailProfissional, DadosProfissional.matricula, DadosProfissional.cpf, ultimoDia, DadosProfissional.reciboBV, DadosProfissional.matricula, DadosProfissional.nome, DadosProfissional.lotacao1, DadosProfissional.regional1, DadosProfissional.categoria, DadosProfissional.dataAdmissao]
  ]

  console.log("Valores a serem enviados para a planilha da manutenção")
  console.log(valores)


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

function envioUnidades(DadosProfissional, ultDia, Respostas){

  let ultimoDia = DadosProfissional.ultimoDia[0] || DadosProfissional.ultimoDia[1] || ultDia
  
  const valores = [
    [DadosProfissional.matricula, DadosProfissional.cpf, DadosProfissional.nome, DadosProfissional.lotacao1, DadosProfissional.regional1, DadosProfissional.categoria, DadosProfissional.dataAdmissao, ultimoDia, DadosProfissional.reciboBV, Respostas["Carimbo de data/hora"], DadosProfissional.emailGestor, DadosProfissional.emailProfissional, Respostas["INICIATIVA DA RESCISÃO"], Respostas["MOTIVOS DA RESCISÃO"]]
  ]

  const arquivos = respostasUnidade.getFiles()

  while(arquivos.hasNext()){
    const arquivo = arquivos.next()
    const nome_arquivo = arquivo.getName()
    const arquivo_regional = nome_arquivo.split("-")[1].trim()
    if(arquivo_regional === DadosProfissional.regional1){
      const arquivo_id = arquivo.getId()
      const plan = SpreadsheetApp.openById(arquivo_id)
      const aba_monitoramento = plan.getSheetByName("Monitoramento de Rescisões - " + arquivo_regional)

      const colunaB = aba_monitoramento.getRange("B3:B").getValues()
      const ultimaLinha = descobrirUltimaLinha(colunaB)
      aba_monitoramento.getRange(ultimaLinha + 1, 2, 1, 14).setValues(valores)
      Logger.log("Enviado à regional " + arquivo_regional)
    }
  }

  
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
/*
function dataLei(){
  const data = new Date(2022,6,1)
  console.log(data)
}
*/