function saudacao(){
  
  const agora = new Date().getHours()
  if(agora < 12){
    return "Bom dia"
  }else if(agora < 19){
    return "Boa tarde"
  }
  return "Boa noite"
}

const assinatura = `<p style=" color: grey; font-size: small;"> <b>Gerência de Gestão dos Contratos Administrativos Temporários - GGCAT</b><br>Secretaria Municipal de Saúde - SMSA/SUS-BH<br>Avenida Afonso Pena, nº 2336, 7º Andar - Funcionários-BH/MG|CEP: 30.130-012<br>(31) 3277-5385 / www.pbh.gov.br</p><img src="cid:logo" alt="PBHlogo" width="150"><p style="font-size:x-small;color: rgb(71, 71, 71);"><i><b>Aviso Legal - Esta mensagem e seus anexos podem conter informações confidenciais e/ou privilegiadas. Se você não for o destinatário ou a pessoa autorizada a recebê-la, não deve usar, copiar ou divulgar as informações nelacontida ou tomar qualquer ação baseada nessas informações, sob pena das ações administrativas, cíveis e penaiscabíveis. Caso entenda ter recebido esta mensagem por engano, por favor, apague-a, bem como seus anexos, e aviseimediatamente ao remetente. Este ambiente é monitorado. A Prefeitura de Belo Horizonte (PBH) informa fazer usopleno do seu direito de arquivar e auditar, a qualquer tempo, as mensagens eletrônicas e anexos processados emseus sistemas e propriedades, com esta declaração eliminando, de forma explícita, clara e completa, qualquer expectativa de privacidade por parte do remetente e destinatários. Decreto Municipal nº 15.423/13</b></i></p> </body></html>`

//----------------------------------------------------------------------------------------------------------------------------------------------------------------

function corpoTextoGestor(nome, reciboBV, confirmaBV){

  console.log("reciboBV recebido no corpoTextoGestor: " + reciboBV)

  var htmlDoc = '<html> <head> <meta charset="UTF-8"> </head><style> body{ font-family:Verdana, Geneva, Tahoma, sans-serif; } </style> <body><p>'

  var semReciboBV = ""

  if(confirmaBV === "Não"){
    semReciboBV = ',<br><br>Notamos que não foi fornecido o número de protocolo da declaração de bens e valores de Desligamento por Iniciativa Própria ou por Iniciativa da Administração Pública Municipal. Já enviamos ao e-mail do profissional o link para a emissão da declaração de bens e valores manual.'
  }

  htmlDoc += saudacao() + '!<br><br> Este é o formulário de rescisão de ' + nome + semReciboBV + '<br><br>Caso esta mensagem não faça sentido para você, gentileza desconsiderá-la.<br><br> Atenciosamente,<br><br></p>'

  htmlDoc += assinatura
    
  return htmlDoc
}

//----------------------------------------------------------------------------------------------------------------------------------------------------------------

function corpoTextoProfissional(nome, matricula, data, confirmaBV, interesse, ciencia){

  //HEAD
  var htmlDoc = '<html> <head> <meta charset="UTF-8"> </head><style> body{ font-family:Verdana, Geneva, Tahoma, sans-serif; } </style> <body><p>'


  //INÍCIO
  htmlDoc += saudacao() + ", " + primeiraMaiuscula(nome.split(" ")[0].toLowerCase())  + "!<br><br>Este e-mail é referente à sua rescisão contratual.<br><br>" 

  if(confirmaBV === "Não"){

    htmlDoc += 'É necessário que você acesse e preencha a declaração de bens e valores manual disponível no link abaixo, uma vez que não foi fornecido o número de protocolo da declaração de bens e valores de Desligamento por Iniciativa Própria ou por Iniciativa da Administração Pública Municipal.<br><br><a href="https://prefeitura.pbh.gov.br/sites/default/files/estrutura-de-governo/planejamento/SUGESP/Formul%C3%A1rios%20instru%C3%A7%C3%B5es%20normativas/00604020_declaracaodebensevalores.pdf">Link para a declaração de bens e valores</a><br><br>'
  }

  if(interesse === "Por interesse do profissional"){

    //Exemplo de link gerado: https://docs.google.com/forms/d/e/1FAIpQLScO2xjj7IASByJvVy2hh7gepFpE8FV_eAjr2RsUiL_QmTguDw/viewform?usp=pp_url&entry.354468775=111111&entry.1490292666=2022-08-05

    const linkNovoForm ="https://docs.google.com/forms/d/e/1FAIpQLScO2xjj7IASByJvVy2hh7gepFpE8FV_eAjr2RsUiL_QmTguDw/viewform?usp=pp_url&entry.354468775=" + matricula + "&entry.1490292666=" + formataData(data)

    htmlDoc += `Pedimos${confirmaBV === "Não" ? " também " : " "}` +  "que você preencha a entrevista de desligamento abaixo.<br><br>" + linkNovoForm + "<br><br>"
  } 

  if (ciencia === "Não, profissional indisponível para contato."){

    htmlDoc += "Esta rescisão foi gerada sem a ciência do profissional por impossibilidade de contato.<br><br>"

  }

  htmlDoc += 'Caso esta mensagem não faça sentido para você, gentileza desconsiderá-la.<br><br>Atenciosamente,<br><br></p>'

  htmlDoc += assinatura

   //Logger.log(htmlDoc) 
  return htmlDoc
    
}

//----------------------------------------------------------------------------------------------------------------------------------------------------------------

function primeirosNomes(nome){

  var nome2 = ""

  if (nome.split(" ")[1].length <= 3){

    nome2 = nome.split(" ")[0] + " " + nome.split(" ")[1] + " " + nome.split(" ")[2]
    
  }else{

    nome2 = nome.split(" ")[0] + " " + nome.split(" ")[1]
  }

  return nome2
}
