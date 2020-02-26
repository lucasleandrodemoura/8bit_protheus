#Include "Protheus.ch"
#Include "ApWebSrv.ch"
#Include "TopConn.ch"
  
  
/*Programa quer pergunta uma chave de uma nota fiscal e o local que quer que o XML seja salvo
*irá consultar no sefaz e se a nota for emitida nos últimos 3 meses irá realizar download na pasta escolhida
*@author Lucas Leandro de Moura <lucas@8bit.inf.br>
*/  
user function lm_dnfe()
	
	msginfo("Este programa tem como objetivo baixar um XML de uma nota fiscal de fornecedor com data de emissão menor que 3 meses")
	Private cPerg        :=  padr("lm_dfe",len(SX1->X1_GRUPO)," ")
	ValidPerg()
	if pergunte(cPerg,.T.)
				
		u_zBxXML(mv_par01,mv_par02)
	endif
	
return	

/*
Seguir nomas do NT2014.002_v1.02_WsNFeDistribuicaoDFe
@param cChaveNfe = Chave a ser consultada
@param cDestino = Local que o XML será salvo
Se desejar poderá ser criado parâmetros
LM_BXUF = Código IBGE do estado que será consultado
LM_BXAMB = Ambiente de emissão 1 produção 2 homologação
LM_BXCA = Certificado CA
LM_BXCER = Certificado digital
LM_BXKEY = Chave do certificado
LM_BXPSW = Senha do certificado
LM_BXPRO = TLS
LM_SSLIN = SSL
*/
User Function zBxXML(cChaveNFe,cDestino)
    Local aArea        := GetArea()
    Local lRet      := .T.
    Local cURL        := "https://www1.nfe.fazenda.gov.br/NFeDistribuicaoDFe/NFeDistribuicaoDFe.asmx?WSDL"
    Local cMsg      := ""
    Local oWsdl     := Nil
    Local cMsgRet   := ""
    Local cError    := ""
    Local cWarning  := ""
    Local cXmlGZip  := ""
    Local cArqXML   := ""
    Local cUfAutor  := SuperGetMV("LM_BXUF",  .F., "43")

    Local cTpAmb    := SuperGetMV("LM_BXAMB", .F., "1")
    Local cCNPJEmp  := Alltrim(Alltrim(FWArrFilAtu(FWCodEmp(),FWCodFil())[18]))
    Local lContinua := .T.
    
    
    Private oXmlDocZip
  
    //Instância a classe, setando as parâmetrizações necessárias
    oWsdl := TWsdlManager():New()
    oWsdl:cSSLCACertFile := SuperGetMV("LM_BXCA",  .F., "")
    oWsdl:cSSLCertFile   := SuperGetMV("LM_BXCER", .F., "")
    oWsdl:cSSLKeyFile    := SuperGetMV("LM_BXKEY", .F., "")
    oWsdl:cSSLKeyPwd     := SuperGetMV("LM_BXPSW", .F., "")
    oWsdl:nSSLVersion    := SuperGetMV("LM_BXPRO", .F., "0")
    oWsdl:lSSLInsecure   := SuperGetMV("LM_SSLIN", .F., .T.)
    oWsdl:nTimeout       := 120
  
    //Tenta fazer o Parse da URL
    lRet := oWsdl:ParseURL(cURL)
    If ! lRet 
        ConOut("Erro ParseURL: " + oWsdl:cError)
        lContinua := .F.
    EndIf
      
    //Se for continuar o processamento
    If lContinua
      
        //Tenta definir a operação
        lRet := oWsdl:SetOperation("nfeDistDFeInteresse")
        If ! lRet 
            ConOut("Erro SetOperation: " + oWsdl:cError)
            lContinua := .F.
        EndIf
    EndIf
      
    //Se for continuar
    If lContinua
        //Monta a mensagem que será enviada
        cMsg := '<soapenv:Envelope xmlns:soapenv="http://www.w3.org/2003/05/soap-envelope">'                        + CRLF
        cMsg += '    <soapenv:Header/>'                                                                              + CRLF
        cMsg += '    <soapenv:Body>'                                                                                + CRLF
        cMsg += '        <nfeDistDFeInteresse xmlns="http://www.portalfiscal.inf.br/nfe/wsdl/NFeDistribuicaoDFe">'  + CRLF
        cMsg += '                <nfeDadosMsg>'                                                                      + CRLF
        cMsg += '                    <distDFeInt xmlns="http://www.portalfiscal.inf.br/nfe" versao="1.01">'          + CRLF
        cMsg += '                        <tpAmb>'+cTpAmb+'</tpAmb>'                                                  + CRLF
        cMsg += '                        <cUFAutor>'+cUfAutor+'</cUFAutor>'                                              + CRLF
        cMsg += '                        <CNPJ>'+cCNPJEmp+'</CNPJ>'                                                  + CRLF
        cMsg += '                      <consChNFe>'                                                                    + CRLF
        cMsg += '                           <chNFe>'+alltrim(cChaveNFe)+'</chNFe>'                                  + CRLF
        cMsg += '                      </consChNFe>'                                                                   + CRLF
        cMsg += '                    </distDFeInt>'                                                                  + CRLF
        cMsg += '                </nfeDadosMsg>'                                                                     + CRLF
        cMsg += '            </nfeDistDFeInteresse>'                                                                 + CRLF
        cMsg += '        </soapenv:Body>'                                                                            + CRLF
        cMsg += '    </soapenv:Envelope>'                                                                            + CRLF
          
        
        //Envia uma mensagem SOAP personalizada ao servidor
        lRet := oWsdl:SendSoapMsg(cMsg)
        If ! lRet 
            ConOut("Erro SendSoapMsg: " + oWsdl:cError)
            ConOut("Erro SendSoapMsg FaultCode: " + oWsdl:cFaultCode)
            lContinua := .F.
        EndIf
    EndIf
  
    //Se for continuar
    If lContinua
        //Pega a resposta do SOAP
        cMsgRet := oWsdl:GetSoapResponse()
                 
        //Transforma a resposta em um objeto
        oXmlDocZip := XmlParser(cMsgRet, "_", @cError, @cWarning)
          
        //Se existir Warning, mostra no console.log
        If ! Empty(cWarning)
            ConOut("Alerta cWarning: " + cWarning)
        EndIf
          
        //Se houve erro, não permitirá prosseguir
        If ! Empty(cError)
            ConOut("Erro cError: " + cError)
            lContinua := .F.
        EndIf
    EndIf
  
    //Se for continuar
    If lContinua
        //Se a tag DocZip existir (for diferente de Undefinied)
        If (Type("oXmlDocZip:_SOAP_ENVELOPE:_SOAP_BODY:_NFEDISTDFEINTERESSERESPONSE:_NFEDISTDFEINTERESSERESULT:_RETDISTDFEINT:_LOTEDISTDFEINT:_DOCZIP") != "U")
  
            //Pega tag que contém XML em zip
            cXmlGZip := oXmlDocZip:_SOAP_ENVELOPE:_SOAP_BODY:_NFEDISTDFEINTERESSERESPONSE:_NFEDISTDFEINTERESSERESULT:_RETDISTDFEINT:_LOTEDISTDFEINT:_DOCZIP:TEXT
              
            //Gera arquivo XML
            cArqXML := fGeraXML(cXmlGZip, cChaveNFe,cDestino)
        else
            msgalert("Ocorreu algum problema no momento de baixar o arquivo da sefaz!")
        endif
    EndIf
  
    RestArea(aArea)
Return cArqXML
  
/*-------------------------------------------------------------------------------*
 Apenas salva o XML no local desejado
 *-------------------------------------------------------------------------------*/  
Static Function fGeraXML(cConteudo, cChave,cDestino)
    Local aArea        := getArea()
    Local lRet         := .T. 
    Local cArquivo     := cChave + ".xml"
    Local cDiretorio := Alltrim(cDestino)
    Local nTamanho     := 0
    Local cUnXML       := ""
    Local cDecode64  := ""
    Local cArqFull   := ""
    Local lHouveErro := .F.
      
    //Se o último caracter do diretório não for \, será barra \
    If SubStr(cDiretorio, Len(cDiretorio), 1) != "\"
        cDiretorio += "\"
    EndIf
      
    //Define o caminho final do arquivo
    cArqFull := cDiretorio + cArquivo
      
    //Pega o tamanho e descriptografa o conteúdo
    nTamanho  := Len(cConteudo)
    cDecode64 := Decode64(cConteudo)
    lRet      := GzStrDecomp(cDecode64, nTamanho, @cUnXML)
       
    //Se deu certo
    If lRet
          
        //Se o diretório não existir, cria
        If ! ExistDir(cDiretorio)
            MakeDir(cDiretorio)
        EndIf
          
        //Cria o arquivo com o conteúdo
        lRet := MemoWrite(cDiretorio+cArquivo, cUnXML)
          
        //Se houve falha, mostra mensagem no console.log
        If ! lRet
            ConOut("Não foi possivel criar o arquivo: " + cArqFull)
            lHouveErro := .T.
        EndIf
      
    //Se não deu certo, mostra mensagem no console.log
    Else
        ConOut("Houve algum erro na descompactação do arquivo!")
        lHouveErro := .T.
    EndIf
  
    //Se houve erro, zera o nome do arquivo para retornar em branco
    If lHouveErro
        cArqFull := ""
    EndIf
  
    RestArea(aArea)
Return cArqFull



//Perguntas
Static Function ValidPerg

	Local cAlias := Alias()
	Local aRegs := {}
	Local i,j

	// Grupo//Pergunta/Variavel/Tipo/Tamanho/Decimal/Presel/GSC/Valid/Var01/Def01/Cnt01/Var02/Def02/Cnt02/Var03/Def03/Cnt03/Var04/Def04/Cnt04/Var05/Def05/Cnt05
	AADD(aRegs,{cPerg,"01","Chave NFE ","","","mv_ch1","C", 44,0,0,"G","","mv_par01","","","","","","","","","","","","","","","","","","","","","","","","","",""})
	AADD(aRegs,{cPerg,"03","Salvar em:","","","mv_ch2","C", 44,0,0,"G","","mv_par02","","","","","","","","","","","","","","","","","","","","","","","","","",""})

	DbSelectArea("SX1")
	DbSetOrder(1)
	//cPerg := PADR(cPerg,6)
	For i:=1 to Len(aRegs)
		If !DbSeek(cPerg+aRegs[i,2])
			RecLock("SX1",.T.)
			For j:=1 to FCount()
				If j<=Len(aRegs[i])
					FieldPut(j,aRegs[i,j])
				Endif
			Next
			MsUnlock()
		Endif
	Next

	DbSelectArea(cAlias)
Return
