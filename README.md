# XL2Csv
Convertir un fichier Excel en fichiers Csv (ou en un fichier txt)
---

XL2Csv est un utilitaire pour convertir directement un fichier Excel en fichiers csv (un fichier csv par feuille Excel, s'il y en a plusieurs), ou bien en un unique fichier csv ou txt.

## Table des matières
- [Utilisation](#utilisation)
- [Techniques](#techniques)
- [Versions](#versions)
- [Liens](#liens)

## Utilisation
Lancez simplement XL2Csv en mode administrateur, cliquez sur OK puis sur "Ajouter menu ctx.".
Ensuite il suffit d'utiliser les menus qui apparaissent lorsque l'on sélectionne un fichier Excel avec le bouton droit de la souris dans l'explorateur de fichiers de Windows :
- Convertir en fichiers Csv (via NPOI)
- Convertir en fichiers Csv (via XLLib.)
- Convertir en un fichier Csv fusionné
- Convertir en un fichier Texte

Pour désinstaller (ou réinstaller XL2Csv en cas de déplacement), cliquez sur "Enlever menu ctx.".

Quelle est la meilleure technologie pour convertir un classeur Excel en csv, comment choisir entre ExcelLibrary et NPOI ?
- ExcelLibrary est la meilleure librairie, elle est rapide. Cependant le code source n'est plus maintenu, il est archivé : https://code.google.com/archive/p/excellibrary
Le package nuget est toujours disponible ici :
https://www.nuget.org/packages/ExcelLibrary
- NPOI fonctionne bien aussi, son seul inconvénient est qu'il indique les bolléens True et False en majuscule (on pourrait les remplacer automatiquement), et son code source est toujours maintenu : https://github.com/nissl-lab/npoi et le package est ici https://www.nuget.org/packages/NPOI

## Techniques
### ODBC (XL2CsvODBC) : Convertir un fichier Excel en fichiers Csv via ODBC

Il s'agit de l'ancienne technique de conversion via ODBC (XL2Csv versions 1.03 et antérieurs) : cette technique est rapide, mais les données doivent être homogènes dans une même colonne (même type de donnée dans une colonne) et la colonne doit avoir un entête. Sinon, des valeurs nulles apparaîtrons à la place des données qui sont de types différents de celui qui a pu être détecté par l'analyse d'un certain nombre de valeurs dans la colonne (1024 valeurs analysées par exemple, pour pouvoir déterminer le type de données de la colonne. Par défaut, seulement 8 valeurs sont analysées, ce qui est insuffisant : en modifiant une clé dans la base de registre, on change en 1024, cela est fait automatiquement par XL2Csv, il faut les droits administrateurs pour ce changement).
Depuis la version 1.04, la méthode par défaut est maintenant la librairie ExcelLibrary, voir ci-dessous.

### Automation (XL2CsvAutomation) : Convertir un fichier Excel en fichiers Csv via Automation Excel

Cette technique utilise l'automation Excel pour lire les valeurs de chaque cellule, en les parcourant une par une, ce qui est très lent (d'autant plus qu'il faut instancier Excel pour cela, ce qui prend déjà du temps). Le seul intérêt de cette technique est pour comparer les résultats avec la librairie utilisée par défaut : ExcelLibrary.

### Fusion csv (XL2CsvGroup) : Convertir un fichier Excel en fichier Csv via ODBC

Cette technique permet de regrouper les données de plusieurs feuilles Excel (supposées de même nature) dans une seule feuille csv, ce qui est fastidieux à faire à la main. Parfois on répartie les données sur plusieurs feuilles Excel pour dépasser la limite des 65000 lignes de données des versions Excel antérieures à 2007, d'où l'intérêt éventuel de cette option.

### Texte (XL2Txt) : Convertir un fichier Excel en fichier Texte

Cette technique permet d'indexer un fichier Excel dans le but de retrouver rapidement un contenu, en affichant l'ensemble des occurrences trouvées dans le classeur (voir l'utilitaire [VBTextFinder](https://github.com/PatriceDargenton/VBTextFinder), une fonctionnalité similaire a été ajoutée à Excel à partir de la version 2010).

## Versions

Voir le [Changelog.md](Changelog.md)

## Liens

Documentation d'origine complète : [XL2Csv.html](http://patrice.dargenton.free.fr/CodesSources/XL2Csv.html)