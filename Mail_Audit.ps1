#Квартальный проект 4 квартал 2022 Наводим порядок в почте "Большая Стирка" (с) Gordeev Oleg
#Цель проекта - Перевод максимального кол-ва обычных почтовых ящиков в Shared Mailbox 
#
#Перед использованием скрипта необходимо подключить себе консоль Exchange
#
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue
Import-Module ActiveDirectory -ErrorAction SilentlyContinue
Clear-Host

$error.Clear()
#Определяем директорию откуда запускается скрипт
$Path = ($MyInvocation.MyCommand.Definition).Replace(($MyInvocation.MyCommand.Name), "")
#Write-Host $Path
$domain = 'domain.local' #Заменить 
$domain_user = 'enter user' #Заменить
$pass = ConvertTo-SecureString 'enter passwd' -AsPlainText -Force
$cred = new-object -typename System.Management.Automation.PSCredential -argumentlist ($domain + "\" + $domain_user), $pass
$mail = 'mailadmin@domain.local' #Заменить
#Подключаемся к серверу Exchange
try{
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exch.domain.local/PowerShell/ -Authentication Kerberos -Credential $cred #Заменить
Import-PSSession $Session -DisableNameChecking  -AllowClobber 
}
catch {
$error[0].Exception.GetType() 
$currentdate = Get-Date
"$currentdate Ошибка $error "| Out-File $Path\MailAuditLog.txt -Append
break
}
####################################################################
 #Запускаем таймер 
$StartTime = (Get-Date)
Start-Sleep -Seconds 10
####################################################################
#Получаем список всех ящиков с типом UserMailbox
$Delegated= 
Get-Mailbox -ResultSize Unlimited | Where-Object { $_.RecipientTypeDetails -like "UserMailbox"} |
#Проверяем какие из них делегированы (исключаем NT AUTHORITY\SELF и группу LSG-Lukinseg-Exchange-Admin-RU) 
Get-MailboxPermission  | Where-Object { $_.user.tostring() -ne "NT AUTHORITY\SELF" -and $_.user.tostring() -notlike "KIFR-RU\LSG-Lukinseg-Exchange-Admin-RU" -and $_.user.tostring() -like "KIFR-RU*" -and $_.IsInherited -eq $false } |
Select-Object Identity | Sort-Object Identity |
#Оставляем уникальные записи т.к. командлет Get-MailboxPermission в выводе перечисляет всех, кому ящик делегирован
Get-Unique -AsString 
#Получаем адреса почты для последующей проверки
$DelegatedMailbox = foreach($a in ($delegated).Identity)
{ Get-Mailbox -Identity $a | Select-Object PrimarySmtpAddress, Alias, database } 
#Write-Host $DelegatedMailbox
#Проверяем список ранее полученных ящиков на предмет отправки и получения ими писем за последние 90 дней 
$Send_log = @()   
foreach ($mailboxitem in $DelegatedMailbox) {
    $Database = Get-MailboxDatabase -Identity $mailboxitem.database
    #$Url = 'http://' + $Database.Server + '.domain.local/PowerShell/'
    
    $ListMessageSend = @()
    $ListMessageReceive = @()
        
        $ListMessageSend +=  Get-MessageTrackingLog  -Server $Database.Server -Sender $mailboxitem.PrimarySmtpAddress -Start (Get-Date).AddDays(-90) -EventID "DELIVER" -resultsize 3 
        $ListMessageReceive += Get-MessageTrackingLog -Server $Database.Server -Recipients $mailboxitem.PrimarySmtpAddress -Start (Get-Date).AddDays(-90) -EventID "RECEIVE" -resultsize 3
   
    $TotalSend = $ListMessageSend | Measure-Object 
    $TotalReceive = $ListMessageReceive | Measure-Object 

   
    $a = $mailboxitem.Alias, $TotalReceive.Count, $TotalSend.Count
    $Send_log += ( , $a)

 
}
#Нет отправок и получений за последние 90 дней - можно отключать 
$NotUsedMailboxes = $Send_log | Where-Object { $_[1] -eq 0 -and $_[2] -eq 0 } 
$temp =foreach ($mailboxitem in $NotUsedMailboxes) {
    Get-ADUser -Identity $mailboxitem[0] -Properties mail | Select-Object Name, SamAccountName, mail
} 
$temp | Export-Csv $Path\NotUsedMailboxes.csv
#Скрываем неиспользуемые ящики 
#foreach($mailboxitem in $NotUsedMailboxes){
#    Set-Mailbox -Identity $NotUsedMailboxes[0] -HiddenFromAddressListsEnabled $true
#}

#Нет отправок, но были получения  
$NotSendButReceive = $Send_log | Where-Object { $_[1] -ne 0 -and $_[2] -eq 0 }  
#Есть отправки
$Sender = $Send_log | Where-Object { $_[1] -ne 0}

#Проверка на роботов 

$Shared = foreach($mailboxitem in $NotSendButReceive) {
    Get-ADUser -Identity $mailboxitem[0] -Properties logoncount -ErrorAction Ignore |
    Where-Object { $_.logoncount -cgt 0 } | Select-Object Name, SamAccountName, logoncount, mail
}

$Check = foreach ($mailboxitem in $Sender) {
    Get-ADUser -Identity $mailboxitem[0] -Properties logoncount -ErrorAction Ignore |
    Where-Object { $_.logoncount -cgt 0 } | Select-Object Name, SamAccountName, logoncount, mail
}
#########################
#Перевод ящиков в Shared#
#########################
#foreach($mailboxitem in $Shared){
#    Set-Mailbox -Identity $Shared.SamAccountName -Type Shared
#}

############################
#Работа с группами рассылок#
############################

#Проверка групп рассылки на предмет пустых групп отключаем 
Get-DistributionGroup –ResultSize Unlimited |
Where-Object { 
    (Get-DistributionGroupMember –Identity $_.Name –ResultSize Unlimited).Count -eq 0
} | 
select Name, PrimarySmtpAddress | Export-Csv $Path\EmptyGroups.csv
#Поиск пустых динамических групп 
Get-DynamicDistributionGroup -ResultSize Unlimited | 
Where-Object { 
    (Get-Recipient -ResultSize 10 -RecipientPreviewFilter (Get-DynamicDistributionGroup -Identity $_.Identity).RecipientFilter).count -eq 0
} | 
select Name, PrimarySmtpAddress | Export-Csv $Path\EmptyDynamic.csv
#Поиск неиспользуемых групп рассылки 
Get-DistributionGroup -ResultSize Unlimited | Select-Object PrimarySMTPAddress | Sort-Object PrimarySMTPAddress | Export-CSV $Path\all_groups.csv 
$Transport = Get-TransportService
$TransportLog = @()
foreach($name in ($Transport).Name){
$GL = Get-MessageTrackingLog -Start (Get-Date).AddDays(-90) -Server $name -EventId Expand -ResultSize Unlimited | Sort-Object RelatedRecipientAddress | 
Group-Object RelatedRecipientAddress | Sort-Object Name | Select-Object @{label = ”PrimarySmtpAddress”; expression = { $_.Name } }, Count | Where-Object {$_.Name -notlike '*_dynamic*'}
$TransportLog += $GL
}
$TransportLog | Export-CSV $Path\active_groups.csv 
$alldl = Import-CSV $Path\all_groups.csv
$activedl = Import-CSV $Path\active_groups.csv
Compare-Object $alldl $activedl -Property PrimarySmtpAddress -SyncWindow 500 | Sort-Object PrimarySmtpAddress | Select-Object -Property PrimarySmtpAddress | Export-Csv $Path\NotUsedGroups.csv

#Отправка результатов проверки
#$NotUsedMailboxes |  Export-Csv $Path\NotUsedMailboxes.csv
$Shared |   Export-Csv $Path\shared.csv
$Check |   Export-Csv $Path\check.csv

$body_common_user = "Во вложении спики ящиков и рассылок, на которые следует обратить внимание:
NotUsedMailboxes.csv - ящики, неиспользуемые последние 90 дней 
Shared.csv - Ящики, которые можно переводить в shared
Check.csv - Ящики, с которыми надо разобраться
EmptyGroups.csv - Пустые группы рассылки
EmptyDynamic.csv - пустые динамические группы
" 
Send-MailMessage -To $mail -From "robot@domain.local" -Subject "Результат аудита почтовых ящиков и рассылок" -SmtpServer smtp.domain.local -Body ($body_common_user) -Attachments $Path\NotUsedMailboxes.csv , $Path\shared.csv , $Path\check.csv , $Path\EmptyGroups.csv , $Path\EmptyDynamic.csv, $Path\NotUsedGroups.csv   -Encoding 'UTF8'

$EndTime = (Get-Date)
$TotalTime = $EndTime-$StartTime
$currentdate = Get-Date
$currentdate.ToString() + '  Время выполнения скрипта  ' + $TotalTime.ToString() | Out-File $Path\MailAuditLog.txt -Append