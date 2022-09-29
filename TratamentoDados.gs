//[10.0, 07924341617, 1215998.0, POLIANA VALADARES DA CRUZ ROQUE, TECNICO DE SERVICOS DE SAUDE, TECNICO EM ENFERMAGEM, CENTRO DE SAÚDE SÃO CRISTÓVÃO, NOROESTE, Wed Feb 19 00:00:00 GMT-03:00 2020, 7.924341617E9, Conselho Regional de Enfermagem MG, 1489269.0, MG - Minas Gerais]


function atribuirValores(dadosADMP, respostas){

  Logger.log("Montando Objeto do Profissional")

  //dados do admp
  const matricula = dadosADMP[3].toString()
  const cpf = dadosADMP[1].toString()
  const nome = dadosADMP[5]
  const categoria = dadosADMP[7].toString()
  const especialidade = dadosADMP[8].toString()
  const numRegistro = dadosADMP[12]
  const conselho = arrumarConselho(dadosADMP[11].toString().toUpperCase()) 
  const ufConselho = dadosADMP[13].toString()
  const dataAdmissao =  dadosADMP[10]
  const lotacao1 = dadosADMP[9].toString()
  const regional1 = dadosADMP[14].toString()



  //dados da resposta do formulario
  let ultimoDia = respostas['DATA DO ÚLTIMO DIA TRABALHADO'].map(data => {
    if(data !== ""){
      return data
    }
  })
  const inicioAviso = respostas['DATA DO INÍCIO DO AVISO PRÉVIO'].toString()
  const comunicacaoAviso = respostas["DATA DA COMUNICAÇÃO DO DIREITO AO AVISO PRÉVIO"].toString()
  const confirmaBV = respostas['DECLARAÇÃO DE BENS E VALORES DE DESLIGAMENTO (DBV)'].toString()
  const reciboBV = respostas['INSIRA O CÓDIGO DO RECIBO DA DBV'].toString()
  const iniciativaRecisao = respostas['INICIATIVA DA RESCISÃO'].toString()
  const avisoPrevio = respostas['O PROFISSIONAL CUMPRIRÁ AVISO PRÉVIO?'].toString()
  const emailGestor = respostas['E-MAIL DO GESTOR'][0].toString() || respostas['E-MAIL DO GESTOR'][1].toString()
  const emailProfissional = respostas['E-MAIL DO PROFISSIONAL'][0].toString() || respostas['E-MAIL DO PROFISSIONAL'][1].toString()

  /*if(emailProfissional === ""){
    emailProfissional = respostas['E-MAIL DO PROFISSIONAL'][1].toString()
  }*/

  //unificando dados
  const Dados = {
    matricula, 
    cpf, 
    nome, 
    categoria, 
    especialidade, 
    numRegistro, 
    conselho, 
    ufConselho, 
    dataAdmissao, 
    lotacao1, 
    regional1, 
    ultimoDia,
    confirmaBV, 
    reciboBV, 
    iniciativaRecisao, 
    avisoPrevio,
    inicioAviso,
    comunicacaoAviso,
    emailGestor, 
    emailProfissional
  }
  
  Logger.log("DadosProfissional:")
  Logger.log(JSON.stringify(Dados))
  return Dados
}

function formataData(Data){

  //Exemplo da url com a data: https://docs.google.com/forms/d/e/1FAIpQLScO2xjj7IASByJvVy2hh7gepFpE8FV_eAjr2RsUiL_QmTguDw/viewform?usp=pp_url&entry.354468775=1234567&entry.1490292666=2022-08-10

  /*const dataArr = Data.split('/')
  const data = new Date(dataArr[2],dataArr[1] - 1,dataArr[0])*/
  const dataFinal = Utilities.formatDate(Data,"GMT -0300", "yyyy-MM-dd")
   
  return dataFinal
}

function hoje(){

  var hj = new Date()

  const dia = hj.getDate() < 10 ? '0' + hj.getDate() : hj.getDate()
  const mes = hj.getMonth()+1 < 10 ? '0' + (hj.getMonth()+1) : (hj.getMonth()+1)
  const ano = hj.getFullYear()

  hj = `${dia}/${mes}/${ano}`
  return hj
  
}

function primeiraMaiuscula(string){

  var pLetra = string.charAt(0).toUpperCase()
  var restoString = string.slice(1)

  return pLetra + restoString
}

function escolheEmail(arrayEmails){
  
}

function arrumarConselho(conselho){

  const soOconselho = conselho.split(" - ")[1]
  return soOconselho
}

//NÃO IMPLEMENTADO
function consertaConselho(conselho){

  const indexUltimaPalavra = conselho.length - 1

  var conselhoArrumado = ""

  const tamanhoUltimaPalavra = conselho.split(" ")[indexUltimaPalavra].length
  
  if(tamanhoUltimaPalavra <= 2){

    conselhoArrumado -= conselho.split(" ")[indexUltimaPalavra]
    return conselhoArrumado
  }
  return conselho
}

