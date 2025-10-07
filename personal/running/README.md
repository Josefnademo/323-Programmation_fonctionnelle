# Exercice : MeteoSport

![running.png](running.png)

Vous allez travailler avec des données météorologiques pour déterminer les meilleurs créneaux horaires pour pratiquer
une activité sportive en plein air...

## Pré-requis

Avant de commencer, récupérez les données météorologiques horaires depuis l'API Open-Meteo. La
structure des données reste la même.

---

## Étape 1 : Récupérer et parser les données

1. Récupérez les données météorologiques depuis l'API Open-Meteo.
2. Désérialisez-les dans des classes C# immuables.

<details>

<summary>Voir une suggestion de solution</summary>

```csharp
using System;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json;

public record WeatherData(Hourly hourly);

public record Hourly(string[] time, float[] temperature_2m, float[] precipitation, float[] wind_speed_10m);

public class Program
{
    public static async Task Main(string[] args)
    {
        using HttpClient client = new HttpClient();
        string url = "https://api.open-meteo.com/v1/forecast?latitude=46.3833&longitude=6.2348&hourly=temperature_2m,precipitation,wind_speed_10m";

        var response = await client.GetStringAsync(url);
        var weatherData = JsonConvert.DeserializeObject<WeatherData>(response);

        // Utilisez weatherData pour les étapes suivantes...
    }
}
```

> Le type `record` au lieu de `class` a été utilisé pour l'immutabilité, car c’est une classe avec des attributs readonly...

</details>

---

## Étape 2 : Filtrer les créneaux idéaux pour le running

1. Créez une fonction pure `GetBestRunningHours`.
2. Cette fonction doit filtrer les créneaux horaires optimaux pour une session de running :
   - Température entre 18 et 22 degrés
   - Pas de pluie
   - Vent faible (<10 kmh)

<details>

<summary>Voir une suggestion de solution</summary>

```csharp
public static IEnumerable<(string Time, float Temperature, float WindSpeed)> GetBestRunningHours(Hourly hourly)
{
    return hourly.time
        .Select((t, index) => new { Time = t, Temperature = hourly.temperature_2m[index], Precipitation = hourly.precipitation[index], WindSpeed = hourly.wind_speed_10m[index] })
        .Where(x => x.Temperature >= 18 && x.Temperature <= 22 && x.Precipitation == 0 && x.WindSpeed < 10)
        .Select(x => (x.Time, x.Temperature, x.WindSpeed));
}
```
</details>

---

## Étape 3 : Calculer les moyennes

1. Créez une fonction pure `CalculateAverages` pour calculer la moyenne de la température et de la vitesse du vent.
2. La fonction doit être immuable et n'altérer aucun des paramètres.

<details>

<summary>Voir une suggestion de solution</summary>


```csharp

public static (float AverageTemperature, float AverageWindSpeed) CalculateAverages(Hourly hourly)
{
    var averageTemperature = hourly.temperature_2m.Average();
    var averageWindSpeed = hourly.wind_speed_10m.Average();

    return (averageTemperature, averageWindSpeed);
}
```
</details>

---

## Étape 4 : Utiliser `Aggregate` pour compter les heures idéales de sport

Créez une fonction pure `CountIdealSportHours` qui utilise `Aggregate` pour compter le nombre d'heures où les conditions sont idéales pour le sport en plein air. Une heure est considérée comme idéale si :
- La température est comprise entre 15°C et 25°C.
- Il n'y a pas de précipitation.
- La vitesse du vent est inférieure à 15 km/h.

<details>

<summary>Voir une suggestion de solution</summary>

```csharp
using System.Linq;

public static int CountIdealSportHours(Hourly hourly)
{
    return hourly.time
        .Select((t, index) => new { Temperature = hourly.temperature_2m[index], Precipitation = hourly.precipitation[index], WindSpeed = hourly.wind_speed_10m[index] })
        .Aggregate(0, (count, hour) =>
        {
            // Incrémenter le compteur si les conditions sont idéales
            if (hour.Temperature >= 15 && hour.Temperature <= 25 && hour.Precipitation == 0 && hour.WindSpeed < 15)
                return count + 1;
            else
                return count;
        });
}


```
</details>

## Étape finale : Afficher les résultats

1. Affichez les résultats des différentes analyses.
2. Utilisez les fonctions pures précédentes pour obtenir les données transformées.

