# excel2qif
Un petit outil pour convertir au format qif les fichiers Excel de ING et les fichiers csv de la SOGE.


# Installation
Pour faciliter les choses, utilisez l'installeur ici :
https://github.com/caliloo/excel2qif/blob/master/installer/excel2qif.exe 

(cliquez sur le lien et ensuite sur le bouton download, il se peut que votre antivirus se plaigne car il est d'origine inconnue, en effet je n'ai pas signé le paquet d'installation, la flemme ... faites simplement un scan, pour vous rassurez, puis lancez le.)

Pour désinstaller, passez par la liste des programmes windows.

# Utilisation
Cet outil ne nécessite pas Microsoft Office.

Pour utiliser le logiciel, faites un clic droit sur un fichier excel ou csv (.xls, .csv), et choisissez convertir au format QIF.
Vous pouvez essayer de faire cette action sur d'autres fichiers excel ou csv que ceux de ING ou de la SOGE, mais le résultat sera alors un fichier QIF au contenu aléatoire (toutefois, votre pc ne devrait pas exploser).

# Documentation
Pour ING, cet outil créé un nouveau fichier avec le même nom, à l'exception de l'extension, qui est renommée en .qif. Le contenu est converti à la volée au format QIF. Si le fichier qif de destination existe déjà, il est alors écrasé par le nouveau fichier QIF.
Pour la SOGE, cet outil créé un fichier dont le nom est au format AAAA MM DD CCCCCCCCCCCC.qif, C étant le numéro de compte récupéré dans l'entête du csv.

Cet outil est fourni avec un installateur pour Windows, mais il devrait fonctioner sans problème sous toute machine capable de faire tourner Python3.7 (vous pouvez installer vous même python3.7 et utiliser la ligne de commande). Il fonctione pour moi (sous windows 10), j'espère sincèrement qu'il fonctionera pour vous. Mais je ne fournis aucune garantie de fonctionnement correct. Vous utilisez cet outil sous votre propre responsabilité, ne lui accordez pas une confiance aveugle, gardez tout de même un oeil sur les résultats, au moins les premières fois....

# Support
Je ne promet aucune réponse rapide ou même tout court, mais je vais regarder un peu cela pendant qqs jours, puis une ou deux fois par semaine. Utilisez svp l'interface de github pour faire remonter les problèmes (ou les choses positives, hein, ca fait plaisir aussi...).
