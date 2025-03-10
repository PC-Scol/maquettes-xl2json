# Objet
Script Python permettant de lire des maquettes de formation dans un fichier Excel (ou csv/texte) et de les convertir dans le format JSON utilisé par la fonction d'importation de maquettes de Pégase. Données injectables dans cette version :
<ul>
  <li>Descripteurs des objets de maquette</li>
  <li>Les champs Enquêtes</li>
  <li>Les champs Syllabus</li>
  <li>Les champs Formats d'enseignement</li>
</ul>

Il est également possible de spécifier le caractère PIA d'un élément, ainsi que préciser le type de relation d'un élément à son parent (obligatoire ou non).

Précision sur la fonction d'import de maquettes de Pégase : cette dernière ne permet pas la mise à jour d'objets déjà présents dans l'environnement où se fait l'import, les objets déjà présents dans Pégase sont simplement pris en lieu et place de ceux décrits dans l'Excel d'import.

<p>&nbsp;</p>

> [!NOTE]
> Notes de mise à jour :
> - Script testé sur la R27 de Pégase
> - Ajout de la possibilité d'injecter le type de liaison d'un élément à son parent (obligatoire/facultatif)
> - Ajout des nouveaux champs d'injection relatifs aux informations Fresq
> - Ajout des champs d'injection relatifs aux formats d'enseignements (sauf les formations porteuses)
> - Ajout de la possibilité d'adapter, dans les Excels de maquettes, les libellés des colonnes, via un fichier d'entêtes (voir l'exemple fourni ici : <i>Entetes-exemple.xlsx</i>)
> Pour tous les points ci-dessus, n'hésitez pas à vous référer au fichier explicatif maquette-type.xlsx.

<p>&nbsp;</p>

> [!TIP]
> J'attire enfin votre attention sur une extension <b>TRES INTERESSANTE</b>, développée par <b><a href="https://github.com/kaisersly">Sylvain Kieffer</a><b>, qui peut contrôler le déclenchement du script maquettes-xl2json.py dans un navigateur Web : cette extension permet donc de rendre autonomes les fonctionnels pour les injections. C'est partagé dans le github de PC-Scol : <b><a href="https://github.com/PC-Scol/maquettes-xl2web">maquettes-xl2web</a></b>.

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
| -e | Spécifie un fichier de définition des entêtes à prendre en compte |
| -n | Spécifie la liste des maquettes à renvoyer au format JSON (séparées par des virgules).<br>Si l'option n'est pas présente, toutes les racines (ie les noeuds de maquette sans parents) sont renvoyées |
| -b | Convertit chaque maquette du flux de sortie en base64 |
| -l | Libellés non obligatoires (un libellé type sera généré automatiquement pour l'import)
| -d | Affiche des messages d'information pour suivre le déroulé de l'execution de la commande |
| -g | Vérifie que les objets de type GROUPEMENT ont bien une plage de choix spécifiée, le script échoue si ce n'est pas le cas
| -c | Utilitaire : renvoie la liste des codes trouvés en entrée (fichiers Excel, textes ou entrée standard)

<p>&nbsp;</p>

# Description d'une maquette type
Le format Excel des maquettes est flexible. Les colonnes attendues obligatoirement sont :
<ul>
  <li>La colonne <bold><i>Type objet</i></bold></li>
  <li>La colonne <bold><i>Code objet</i></bold></li>
  <li>La colonne <bold><i>Libellé </i></bold> (celle-ci pouvant être rendue optionnelle via l'option de commande <bold><i>-l</i></bold>)</li>
</ul>

Toutes les autres colonnes sont optionnelles.<br>
Par ailleurs, l'ordre des colonnes n'est pas figé et d'autres colonnes (commentaires, formules, zones de validation, bref tout ce qu'Excel sait si bien faire) peuvent être insérées là où cela vous est nécessaire. La seule contrainte est de garder le bon entête de colonne (le libellé sur fond bleu dans la maquette type).

<p>&nbsp;</p>

# Couplage avec le script d'upload de maquettes
En chaînant le script maquettes-xl2json.py avec le script d'upload vers Pégase maquettes-upload.py, on obtient une fonctionnalité d'upload direct de maquettes au format Excel vers Pégase.

Par exemple :
```bash
  maquettes-xl2json.py maquette-type.xslx | maquettes-upload.py inalco BAS ESPACE-TEST
```
pour téléverser dans Pégase la maquette du fichier maquette-type.xlsx vers l'instance bac à sable (BAS) de l'Inalco, dans l'espace de travail ESPACE-TEST.
