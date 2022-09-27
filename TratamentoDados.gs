//[10.0, 07924341617, 1215998.0, POLIANA VALADARES DA CRUZ ROQUE, TECNICO DE SERVICOS DE SAUDE, TECNICO EM ENFERMAGEM, CENTRO DE SAÚDE SÃO CRISTÓVÃO, NOROESTE, Wed Feb 19 00:00:00 GMT-03:00 2020, 7.924341617E9, Conselho Regional de Enfermagem MG, 1489269.0, MG - Minas Gerais]


function atribuirValores(dadosADMP, respostas){

  Logger.log("Montando Objeto do Profissional")

  //dados do admp
  const matricula = trataMatricula(dadosADMP[0].toString()) 
  const nome = dadosADMP[1]
  const cpf = trataCPF(dadosADMP[2].toString()) 
  const categoria = dadosADMP[3].toString()
  const especialidade = dadosADMP[4].toString()
  const lotacao1 = dadosADMP[5].toString()
  const dataAdmissao =  dadosADMP[6] 
  const conselho = arrumarConselho(dadosADMP[7].toString().toUpperCase()) 
  const numRegistro = dadosADMP[8]
  const ufConselho = dadosADMP[9].toString()
  const regional1 = dadosADMP[10].toString()


  //dados da resposta do formulario
  const ultimoDia = respostas['DATA DO ÚLTIMO DIA TRABALHADO'].toString()
  const confirmaBV = respostas['DECLARAÇÃO DE BENS E VALORES DE DESLIGAMENTO (DBV)'].toString()
  const reciboBV = respostas['INSIRA O CÓDIGO DO RECIBO DA DBV'].toString()
  const iniciativaRecisao = respostas['INICIATIVA DA RESCISÃO'].toString()
  const emailGestor = respostas['E-MAIL DO GESTOR'].toString()
  var emailProfissional = respostas['E-MAIL DO PROFISSIONAL'][0].toString()

  if(emailProfissional === ""){
    emailProfissional = respostas['E-MAIL DO PROFISSIONAL'][1].toString()
  }

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
    emailGestor, 
    emailProfissional
  }
  
  Logger.log("Objeto profissional montado:")
  Logger.log(JSON.stringify(Dados))
  return Dados
}

function formataData(Data){

  //Exemplo da url com a data: https://docs.google.com/forms/d/e/1FAIpQLScO2xjj7IASByJvVy2hh7gepFpE8FV_eAjr2RsUiL_QmTguDw/viewform?usp=pp_url&entry.354468775=1234567&entry.1490292666=2022-08-10

  const dataArr = Data.split('/')
  const data = new Date(dataArr[2],dataArr[1] - 1,dataArr[0])
  const dataFinal = Utilities.formatDate(data,"GMT -0300", "yyyy-MM-dd")
   
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

function trataMatricula(matr){

  if(matr.length < 7){

    var matrCom0 = matr

    for (let i = 0; i < 7 - matr.length; i++){
      
      matrCom0 = "0" + matrCom0
    }
    return matrCom0
  }
  return matr
}

function trataCPF(CPF){
  
  if(CPF.length < 11){

    var CPFCom0 = CPF

    for (let i = 0; i < 11 - CPF.length; i++){
      
      CPFCom0 = "0" + CPFCom0
    }
    return CPFCom0
  }
  return CPF
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

