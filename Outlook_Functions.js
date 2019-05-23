function exemplo_enviarEmail(){
	// Envio de email por script javascript usado em cscript
	
	// Inicio - NÃO EXCLUIR!
	var tm_total = new Timer();	

	// Variaveis
	var strCaminho = "C:\\temp\\";
	var strAquivo01 = "Aquivo01.pdf";
	var strAquivo02 = "Aquivo02.pdf";

	// Cadastro de operacao
	objOperacoes.push({ "Emails": ["p000ailton.zacarias@teste.com.br","lucas@teste.com.br","leandro@teste.com.br"] ,
		"CC": ["p000ailton.zacarias@teste.com.br","lucas@teste.com.br","leandro@teste.com.br"] ,
		"Anexos": [strCaminho + strAquivo01 , strCaminho + strAquivo02] ,
		"Subject" : "Teste de Subject" ,
		"Body" : "Teste de Body" ,
		"Status" :""
	});

	// Enviar eMail
	for(var i=0; i< objOperacoes.length ; i+=1 ) {
		enviarEmail(objOperacoes[i],modelo);	
	}
	
	// Terminio - NÃO EXCLUIR!
	return tm_total.elapsed();
}

function enviarEmail(obj,modelo) 
{
    var outlook = new ActiveXObject("Outlook.Application");
	
	if( modelo != undefined )
		var emailItem = outlook.CreateItemFromTemplate(modelo);
	else
	{
		var emailItem = outlook.CreateItem(0);
		emailItem.Body = obj.Body;		//Texto / Corpo do eMail [Body]
	}

	with(emailItem){

		var recipienteDestinatarios = recipients();
		var recipienteCC = recipients();

		// Assunto
		Subject = obj.Subject;
		
		// Anexo
		for (var i = 0; i < obj.Anexos.length; i++)
			Attachments.Add(obj.Anexos[i].trim());
		
		// Destinatários
		for (var i = 0; i < obj.Emails.length; i++)		
			if ( obj.Emails[i].trim() != "" ) 
				recipienteDestinatarios.Add(obj.Emails[i].trim());
		
		// Cópia
		for (var i = 0; i < obj.CC.length; i++){
			recipienteCC.Add(obj.CC[i].trim());
			recipienteCC( recipients().count ).type = 2;
		}

		// Confirmação de leitura
		ReadReceiptRequested = true;
		
		//Enviar
		Send();
	}
}

