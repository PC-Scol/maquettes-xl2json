# Objet
Script Python permettant de lire des maquettes de formation dans un fichier Excel (ou csv/texte) et de les convertir dans le format JSON utilisé par la fonction d'importation de maquettes de Pégase.
> [!NOTE]
> Script testé sur chaque version Pégase depuis la V24
> Aucune API n'a été maltraitée lors de la création de ce script.

<p>&nbsp;</p>

# Usage
```bash
  maquettes-xl2json.py fichier_excel ou répertoire
```

- Il est possible de spécifier plusieurs noms de fichiers à la suite, qui seront traités successivement.
- Il est également possible de spécifier un ou plusieurs répertoires : dans ce cas, chacun d'eux sera parcouru récursivement et tout Excel trouvé sera traité.

<p>&nbsp;</p>

> [!TIP]
> Pour déterminer les onglets à traiter dans un fichier Excel, on ajoute la liste des onglets (séparés par __`:`__) après le nom du fichier.
> Par exemple `maquettes-xl2json.py ma_maquette.xlsx:2:3:5` pour traiter les onglets 2, 3 et 5.
> - Si aucun onglet n'est indiqué, c'est le 1er onglet de l'Excel est traité
> - Si on a juste __`:`__ derrière le nom de fichier, tous les onglets sont traités
> - On peut aussi spécifier un nom d'onglet ou un début de nom : les onglets commençant par ce nom seront tous analysés

<p>&nbsp;</p>

Par ailleurs, la commande peut contenir les paramètres suivants :

| Option | Description |
| --- | --- |
| -a | Affiche un message d'aide |
| -n | Spécifie la liste des maquettes à renvoyer au format JSON (séparées par des virgules).<br>Si l'option n'est pas présente, toutes les racines (ie les noeuds de maquette sans parents) sont renvoyées |
| -b | Convertit chaque maquette du flux de sortie en base64 |
| -d | Affiche des messages d'information pour suivre le déroulé de l'execution de la commande |
| -h | En cas de besoin, pour les cas désespérés : transmet les indices de colonnes (séparés par virgule) où se trouvent les données dans l'Excel source. Les indices correspondent aux données suivantes :<br>type d'objet, nature, code, libellé, libellé long, ects, code du noeud parent, plage min de groupement, plage max |

<p>&nbsp;</p>

# Couplage avec le script d'upload de maquettes
En chaînant le script maquettes-xl2json.py avec le script d'upload vers Pégase maquettes-upload.py, on obtient une fonctionnalité d'upload direct de maquettes au format Excel vers Pégase.

Par exemple :
```bash
  maquettes-xl2json.py maquette-type.xslx | maquettes-upload.py inalco BAS ESPACE-TEST
```
pour téléverser dans Pégase la maquette du fichier maquette-type.xlsx vers l'instance bac à sable (BAS) de l'Inalco, dans l'espace de travail ESPACE-TEST.
