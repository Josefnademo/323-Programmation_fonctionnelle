# Immutable

Ces exercices ont pour but de vous apprendre à repérer les variables mutables et à trouver des solutions dans lesquelles on ne se sert que de variables immutables

## 1. Maximum

Dans un programme de jeu de carte qui se joue à 4 joueur, on a décidé que le joueur le plus âgé serait celui qui commence.

On sera donc tenté d'écrire le code suivant:

<details>
<summary>Afficher/Masquer le code</summary>
    
```csharp
// 4 players
List<Player> players = new List<Player>()
{
    new Player("Joe", 32),
    new Player("Jack", 30),
    new Player("William", 37),
    new Player("Averell", 25)
};

// Initialize search
Player elder = players.First();
int biggestAge = elder.Age;

// search
foreach (Player p in players)
{
    if (p.Age > biggestAge) // memorize new elder
    {
        elder = p;
        biggestAge = p.Age; // for future loops
    }
}

Console.WriteLine($"Le plus agé est {elder.Name} qui a {elder.Age} ans");

Console.ReadKey();

public class Player
{
    private readonly string _name;
    private readonly int _age;

    public Player(string name, int age)
    {
        _name = name;
        _age = age;
    }

    public string Name => _name;

    public int Age => _age;
}
```

</details>

### TODO
1. Créer un dossier `maximum`dans votre dossier `personnel` 
2. Créer un projet console dans ce dossier, avec le code ci-dessus
3. Faire un commit nommé `chore(exo maximum): Importer le code de base`
4. Identifier la ou les variables mutables, puis trouver et implémenter une manière de n'avoir que des variables immutables **sans avoir recours à LinQ** (vous pouvez vous appuyer sur le fait que la liste aura toujours exactement 4 joueurs)
5. Faire un commit nommé `feat(exo maximum): Remplacer les variables mutables par des immutables`

---

## 2. Variable globale

Le code ci-dessous compile et s'éxécute mais quelquechose fait que son résultat laisse songeur...

<details>
<summary>Afficher/Masquer le code</summary>

```csharp
using System;
using System.Threading;

class Program
{
    static int linesOfCode = 9210;
    static Random random = new Random();

    static void Main(string[] args)
    {
        Console.WriteLine($"Opening shop with {linesOfCode} lines in our program");
        Thread thread1 = new Thread(Bob);
        Thread thread2 = new Thread(Alice);

        // Start both threads
        thread1.Start();
        Thread.Sleep(300); // Alice starts her day a bit later
        thread2.Start();

        // Wait until both threads terminate
        thread1.Join();
        thread2.Join();

        Console.WriteLine($"Closing shop with {linesOfCode} lines in our program");
        Console.ReadLine();
    }

    static void Bob()
    {
        int myLineCounter = linesOfCode; // Bob starts working
        int workingHours = 0;

        while (workingHours < 9) // He has a 9-hours day ahead of him
        {
            Thread.Sleep(1000); // Bob works for "1 hour"
            workingHours++;
            int BobProduction = random.Next(10, 50);
            Console.WriteLine($"Bob commits {BobProduction} lines of code.");
            myLineCounter += BobProduction;
        }
        Console.WriteLine($"Bob checks out, he claims the program has now {myLineCounter} lines");
        linesOfCode = myLineCounter; // he turns his work in
    }

    // Method to be executed by thread2 - increments by 2 every 3 seconds
    static void Alice()
    {
        int myLineCounter = linesOfCode; // Alice starts working
        int workingHours = 0;
        while (workingHours < 7) // She has a 7-hours day ahead of her
        {
            Thread.Sleep(1000); // Alice works for "1 hour"
            workingHours++;
            int AliceProduction = random.Next(20, 80);
            Console.WriteLine($"Alice commits {AliceProduction} lines of code.");
            myLineCounter += AliceProduction;
        }
        Console.WriteLine($"Alice checks out, she claims the program has now {myLineCounter} lines");
        linesOfCode = myLineCounter; // she turns her work in
    }
}

```

</details>

1. Rappatrier cet échantillon dans un nouveau projet
2. Lancer son éxécution
3. Constater le problème
4. Annoter avec des commentaires partout où il faudrait faire des corrections
5. git stage Program.cs && git commit -m "chore(comments): added comments to fix ..." (compléter les ...)

**Ensuite,**

### Option 1
 - Créer un test d'integration (MSTest, XUNIT,...) piloté par du code qui ne passe pas (valeur finale incorrecte)
 - Appliquer les corrections indiquées et vérifier que le test passe

### Option 2
Bypasser la première étape de l'option 1 et vérifier à la main ensuite

**Puis**
stage/add,commit, push
