# Extrair a lista de todos os usuários do Active Directory
# Criado por: Bruno Farias Novaes
# E-mail: bruno.novaes@competeti.com.br
# WebSite: http://www.competeti.com.br
# Executar: 
# ./ListarUsersAD.ps1

$relatorio = $null
$tabela = $null
$data = Get-Date -format "dd/MM/yyyy"
#$datename = Get-Date -format "yyyy/MM/dd"
#$filename = $datename.Replace("/","-").Replace(":","-")
$arquivo = "SRVAD01.html"
$total = (Get-ADUser -filter *).count
$dominio = (Get-ADDomain).Forest
$analista = "BRUNO FARIAS NOVAES"
$empresa = "COMPETE TI - Serviços e Projetos de TI"

Import-Module ActiveDirectory

#--LISTA DE USUÁRIOS------------------------------------------------#
$tabela += "<center><h2><b>TOTAL DE USUARIOS - <font color=red>$total</font></b></h2></center>"
#Name,SamAccountName,PasswordNeverExpires,Enabled,DistinguishedName,DisplayName,GivenName,SurName,Company,Mail,Department,Title

#$usuarios = @(Get-ADUser -filter * -Properties SamAccountName, Name, DistinguishedName, Department, Enabled, Created)
#$resultado = @($usuarios | Select-Object SamAccountName, Name, DistinguishedName, Department, Enabled, Created)

#$usuarios = @(Get-ADUser -filter * -Properties Name,SamAccountName,PasswordNeverExpires,Enabled,DistinguishedName,DisplayName,GivenName,SurName,Company,Mail,Department,Title)
#$resultado = @($usuarios | Select-Object Name,SamAccountName,PasswordNeverExpires,Enabled,DistinguishedName,DisplayName,GivenName,SurName,Company,Mail,Department,Title)

$usuarios = @(Get-ADUser -filter * -Properties Name, SamAccountName, DistinguishedName, Enabled, Created)
$resultado = @($usuarios | Select-Object Name, SamAccountName, DistinguishedName, Enabled, Created)

$resultado = $resultado | Sort -Property Created -Descending

# Comente esta linha para não exibir o resultado durante a execução do script.
#$resultado | ft -auto 

$tabela += $resultado | ConvertTo-Html -Fragment
 
$formatacao="<html>
		<body>
		<style>
		BODY{font-family: Arial; font-size: 10pt;}
		TABLE{border: 1px solid black; border-collapse: collapse; font-size: 10pt; text-align:center;margin-left:auto;margin-right:auto; width='1600px';}
		TH{border: 1px solid black; background: #F9F9F9; padding: 5px;}
		TD{border: 1px solid black; padding: 5px;}
		H3{font-family: Calibri; font-size: 10pt;}
		</style>"
		
$titulo="<table width='100%' border='0' cellpadding='0' cellspacing='0'>
		<tr>
		<td bgcolor='#F9F9F9'>
		<font face='Calibri' size='5px'>LISTA DE USUARIOS DA REDE</font>
		<H5 align='center'>EMPRESA: $empresa - DOMINIO: $dominio <br> RELATORIO: $data - RESPONSAVEL: $analista</H5>
		</td>
		</tr>
		</table>
		</body>
		</html>"
		
$mensagem = "</table><style>"
$mensagem = $mensagem + "BODY{font-family: Calibri;font-size:20;font-color: #000000}"
$mensagem = $mensagem + "TABLE{margin-left:auto;margin-right:auto;width: 800px;border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$mensagem = $mensagem + "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color: #F9F9F9;text-align:center;}"
$mensagem = $mensagem + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;text-align:center;}"
$mensagem = $mensagem + "</style>"
$mensagem = $mensagem + "<table width='349px'  heigth='400px' align='center'>"
$mensagem = $mensagem + "<tr><td bgcolor='#DDEBF7' height='40'>AUDITORIA</td></tr>"
$mensagem = $mensagem + "<tr><td height='80'>Lista completa de todos os <b>usuarios</b> do Active Directory</td></tr>"
$mensagem = $mensagem + "<tr><td bgcolor='#DDEBF7' height='40'>SEGURAN&#199;A DA INFORMA&#199;&#195;O</td></tr>"
$mensagem = $mensagem + "</table>"

$relatorio = $formatacao + $titulo + $tabela

#--GERAR O HTML-----------------------------------------------------#
$relatorio | Out-File $arquivo -Encoding Utf8

# Exportar para o formato CSV (ad-lista.csv)
$resultado | Sort Name | Export-Csv SRVAD01.csv -NoTypeInformation -Encoding Utf8

#--ENVIAR RELATÓRIO VIA E-MAIL--------------------------------------#
#$de = "bruno.novaes@competeti.com.br"
#$para = "bruno.novaes@competeti"
#$assunto = "USERS AD - $data"
#$smtp = "SMTP_SERVER"
#$porta = "587"
#Send-MailMessage -From $de -To $para -Subject $assunto -Attachments $arquivo,ad-lista.csv -bodyashtml -Body $mensagem -SmtpServer $smtp -Port $porta
