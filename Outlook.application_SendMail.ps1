#
#   Send mail via Outlook with attachment
#
#
#  By Thom Kortekaas
#  Created:  22-11-2013
#  Modified: 22-11-2013
#

##########################################################
### Fill variables                                     ###
##########################################################
$Attachment   = "C:\Users\Thom\Desktop\testfile.txt"
$Recipients   = "user@domain.tld"
$Subject      = "Rapport 2013"
$Body         = "See attachment"  

##########################################################
### Setup mail                                         ###
##########################################################

#Create outlook instance
$Outlook = new-object -comobject outlook.application  
  
#Create new mail item
$NewMail = $Outlook.CreateItem(0)

#Display mail item to user  
$NewMail.display()

##########################################################
#Fill email with data, 1 second delay for each step    ###
##########################################################
#Fill in recipients
start-sleep -s 1
$NewMail.To = $Recipients

#Fill in subject
start-sleep -s 1
$NewMail.Subject = $Subject  

#Fill in body while preserving current body (signature)
start-sleep -s 1
$oldbody = $NewMail.HTMLBody
$newbody  =  $Body
$newbody += $oldbody
$NewMail.HTMLBody = $newbody

##########################################################
### Add attachment if given                            ###
##########################################################
#Check if variable $Attachment has a value to begin with, if not skip attachment
if ($Attachment){
#Check path of given attachment
if(Test-Path $Attachment){$NewMail.Attachments.add($Attachment)}else {
#If attachment can't be found view MsgBox
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
[System.Windows.Forms.MessageBox]::Show("Can't find $Attachment, check your path! ") 

}
}

start-sleep -s 3
#Remove "#" below if mail should be send automatically 
#$NewMail.Send() 