# WordIndexMacros
Microsoft Word Macros to buid an index with the paragraph numbers

# Numéroter les paragraphes et construire un index renvoyant aux numéros de paragraphes

## 1. Numéroter les paragraphes

Ne prendre que le corps du texte : copier-coller dans un nouveau document, enlever l’avant-texte qui ne doit pas être numéroté, et enlever la fin du document (tables, bibliographie, annexes…) qui ne doit pas recevoir de numérotation.
Pressez en même temps les touches Alt et F11.  Cliquez deux fois sur ThisDocument à gauche, et collez la macro **MetLesNumeros()**.

## 2. Indexation

Indexez votre document. Les champs d’index doivent être des champs XE classiques : { XE "Entrée" }.

## 3. Suppression de certains champs d’index

Word indexe les notes de bas de page, les titres, les tables (sommaire et table des matières, etc.).
Si vous préférez ne pas indexer ces éléments, utilisez les macros suivantes.
Pour l’installation des macros, voir ci-dessus le paragraphe 1.

### 3.1. Effacer les champs XE dans les notes
Utilisez la macro **EffaceEntreesIndexDansNotes()**

### 3.2. Effacer les champs XE dans les titres
Utilisez la macro **EffaceEntreesIndexDansTitres()**

### 3.3. Effacer les champs XE dans les tables
Utilisez la macro **EffaceEntreesIndexDansTables()**

## 4. Modifier les champs d'index

Pressez en même temps les touches Alt et F11.  Cliquez deux fois sur ThisDocument à gauche, et collez la macro **InsereRefStyle()**.
Cette macro insère dans chaque champ XE un nouveau champ STYLEREF qui permet de récupérer automatiquement le numéro du paragraphe courant.
Puis cliquez sur le bouton Exécuter (F5)

## 5. Générer l'index

Placez le curseur à l’endroit où l’index prendra place. Dans l’onglet Références, cliquez sur Insérer l’index.
