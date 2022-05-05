$url = "https://winofsql.jp/image/planet.jpg"

Add-Type -path "wget.cs" `
    -ReferencedAssemblies System.Web, System.Windows.Forms

[Program]::Main($url)

Read-Host "‰½‚©ƒL[‚ğ‰Ÿ‚µ‚Ä‚­‚¾‚³‚¢"