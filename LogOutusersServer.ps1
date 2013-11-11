function RemoveSpace([string]$text) {  
    $private:array = $text.Split(" ", `
    [StringSplitOptions]::RemoveEmptyEntries)
    [string]::Join(" ", $array) }
	
$admin = [Environment]::UserName

$sessies = quser
foreach ($sessie in $sessies) {
    $sessie = RemoveSpace($sessie)
    $DEsessie = $sessie.split()
    
    if ($DEsessie[0].Equals(">$admin")) {
    continue }
   if ($DEsessie[0].Equals("Gebruikersnaam")) {
    continue }
 	 c:\windows\system32\logoff.exe $DEsessie[1] }

