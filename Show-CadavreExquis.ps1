<#
.SYNOPSIS
	Cadavre Exquis.

.DESCRIPTION
	Starts the Notepad application, makes it active, and injects characters (Senteces) through an infinite loop.
    Stop with CTRL+C or stop button if start with ISE.

.NOTES

    Author : Vincent Dubois

#>

# Start a Notepad Process
# To Do, need to be alone, or test for the empty one

# & mean another instance
& Start-Process 'C:\WINDOWS\system32\notepad.exe'

# Get list of active process (Id and name)
$process = Get-Process | Where-Object {$_.MainWindowTitle -ne ""} | Select-Object Id, name

# For each $process search the first "notepad", select is Id
foreach ($p in $process) {
    if ($p.Name -eq "notepad") {
        Write-Host $("found ", $p.Name, " ", $p.Id )
        $nId = $p.Id
    }
}

# Notepad activated, window in front and active
$null = (New-Object -ComObject WScript.Shell).AppActivate($nId)

# create a shell for key injection
$myshell = New-Object -com "Wscript.Shell"
# send message and ENTER
$myshell.sendkeys("Boucle infinie !{ENTER}")
$myshell.sendkeys("Stopper le process manuelment.{ENTER}")
$myshell.sendkeys("Création de phrases:{ENTER}{ENTER}")

# Random
#$rand = Get-Random
#Get-Random -SetSeed $rand

# Sentence
$LaPhrase = ""

# charge les articles
$Articles = ("un", "le", "une", "la")

# charge les noms
$NomsMasculins = ("avion", "bateau", "cadavre", "camion", "cretin", "delinquant",
    "elephant", "homme", "policier", "steak", "téléscope", "zombi", "zonard", "mug",
    "cafe", "squelette", "poisson", "marron", "helicopter", "train", "fantome",
    "garcon", "gars", "mec", "clou", "marteau", "ciseaux", "couteau", "chien",
    "chat", "cheval", "pistolet", "revolver")

$NomsFeminins = ("camionette", "courgette", "delinquante", "femme", "grenade",
    "merguez", "policiere", "saucisse", "voiture", "tomate", "pomme", "orange", "tasse",
    "zombie", "piscine", "armoire", "chateigne", "fille", "meuf", "catapulte", "fourchette",
    "cuillere", "sonette", "vipere", "chaussette", "carabine")

# charge les compléments
$Complements = ("bleu.e", "intelligent.e", "ivre", "orange", "rouge", "stupide", "vert.e",
    "blanc.he", "exquis.e", "noir.e", "violet.e", "marron", "debile", "idiot.e")

# charge les verbes
$Verbes = ("allume", "assemble", "bois", "coule", "derange", "écope", "farfouille", "mange",
    "cherche", "pousse", "vol", "attrape", "syphone", "est", "a", "avait", "prend", "mangait",
    "pari sur", "bloque", "catapulte", "enfonce", "defonce", "est aller voir", "a attraper",
    "a mordu", "explose", "tire sur")


# infinite loop
do {
    $LaPhrase = ""
    # article
    $Article = $Articles | Get-Random
    $LaPhrase = -join ($LaPhrase, $Article, " ")
    
    # Nom selon article
    if ( ($Article -eq "une") -or ($Article -eq "la") ) {
        $Nom = $NomsFeminins | Get-Random
    }
    else {
        $Nom = $NomsMasculins | Get-Random
    }
    $LaPhrase = -join($LaPhrase, $Nom, " ")

    # complément
    $Complement = ""
    $Complement = $Complements | Get-Random
    $LaPhrase = -join($LaPhrase, $Complement, " ")

    # verbe
    $Verbe = ""
    $Verbe = $Verbes | Get-Random
    $LaPhrase = -join($LaPhrase, $Verbe, " ")

    # article
    $Article = $Articles | Get-Random
    $LaPhrase = -join($LaPhrase, $Article, " ")

    # nom selon article
    if ( ($Article -eq "une") -or ($Article -eq "la") ) {
        $Nom = $NomsFeminins | Get-Random
    }
    else {
        $Nom = $NomsMasculins | Get-Random
    }
    $LaPhrase = -join($LaPhrase, $Nom, " ")


    # complément
    $Complement = $Complements | Get-Random
    $LaPhrase = -join($LaPhrase, $Complement, ".")

    Start-Sleep -Seconds 5

    # each minute show it
    $myshell.sendkeys("{ENTER}")
    $myshell.sendkeys($LaPhrase)
    $myshell.sendkeys("{ENTER}")

}until(0)

