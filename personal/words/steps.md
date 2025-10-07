# Etapes de réalisation

Si vous savez comment vous attaquer à ce problème, passez votre chemin.

Si vous ne savez pas trop comment vous y prendre, voici des suggestions:

1. Créez une fonction `double Epsilon (string word)` qui retourne une valeur quelconque. Par exemple: `return Math.sqrt(word.length)`
2. Utilisez cette fonction comme critère de filtre
3. Remplacez le contenu de la fonction `Epsilon`:
   - Mettez les données statistiques dans un dictionnaire (`Dictionary<char,double>`)
   - Faites en sorte que la valeur retournée soit la somme des probabilités de toutes les lettres, sans vous soucier de celles qui réapparaissent plusieurs fois dans le mot.
   Si vous n'êtes pas encore à l'aise avec LinQ, écrivez d'abord le code avec ce que vous connaissez, puis transformez-le.
4. Dans la fonction `Epsilon`, créez un dictionnaire (`Dictionary<char,int>`) qui dit pour chaque lettre du mot combien de fois elle apparaît. Utilisez ce dictionnaire pour ajuster la valeur retournée
