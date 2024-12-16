# Deve ser executado direto no servidor que hospeda a role de Certification Authority

# Arquivo com o resultado final
$ResultFile = "C:\temp\CA4.csv"

# Obtém a lista de certificados emitidos pela CA
$issuedCerts = (certutil -out "Request Disposition,Revocation Date,Certificate Expiration Date,Certificate Effective Date,Certificate Template,Issued Common Name,Issued Email Address" -view csv) 

# Filtra o resultado por condições específicas
$filtered = $issuedCerts | ConvertFrom-Csv | Where-Object {$_."Request Disposition" -eq "20 -- Issued" -and $_."Revocation Date" -eq "EMPTY" -and $_."Certificate Template" -ne "CAExchange"}| Where-Object {[DateTime]$_."Certificate Expiration Date" -gt (Get-Date)}

# Formata o cabeçalho do arquivo de resultado
$msg = "Request Disposition;Revocation date;Certificate Expiration Date;Certificate Effective Date;Certificate Template;Issued Common Name;Issued Email Address"
Add-Content -Value $msg -Path $ResultFile
	
# Para cada certificado, formata o resultado e salva no arquivo de resultado
ForEach ($cert in $filtered){
	
	$CertExpireDate = Get-Date($cert."Certificate Expiration Date")
	#$CertExpireDate.ToString("dd/MM/yyyy")
	
	$CertEffectiveDate = Get-Date($cert."Certificate Effective Date")
	#$CertEffectiveDate.ToString("dd/MM/yyyy")
	
	$msg = "" + $cert."Request Disposition" + ";" + $cert."Revocation date" + ";" + $CertExpireDate.ToString("dd/MM/yyyy") + ";" + $CertEffectiveDate.ToString("dd/MM/yyyy") + ";" + $cert."Certificate Template" + ";" + $cert."Issued Common Name" + ";" + $cert."Issued Email Address"
	
	Add-Content -Value $msg -Path $ResultFile
}
