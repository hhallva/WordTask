namespace DocumentsLibrary.Models
{
    public static class Model
    {
        public static List<Genre> Genres = new List<Genre>
        {
            new Genre { Id = 1, Name = "Action" },
            new Genre { Id = 2, Name = "Adventure" },
            new Genre { Id = 3, Name = "RPG" },
            new Genre { Id = 4, Name = "Strategy" },
            new Genre { Id = 5, Name = "Simulation" }
        };

        public static List<Game> Games = new List<Game>
        {
            new Game { Id = 1, Name = "Space Combat 2077", Description = "Fast-paced space action!", GenreId = 1, Genre = Genres[0] },
            new Game { Id = 2, Name = "The Lost City of Eldoria", Description = "Explore ancient ruins in this epic adventure.", GenreId = 2, Genre = Genres[1] },
            new Game { Id = 3, Name = "Swords and Sorcery Online", Description = "A vast open-world RPG with endless possibilities.", GenreId = 3, Genre = Genres[2] },
            new Game { Id = 4, Name = "Galactic Conquest", Description = "Lead your civilization to galactic dominance in this strategy game.", GenreId = 4, Genre = Genres[3] },
            new Game { Id = 5, Name = "Farm Life Simulator", Description = "Experience the joys of rural life.", GenreId = 5, Genre = Genres[4] },
            new Game { Id = 6, Name = "Urban Assault: Reloaded", Description = "Intense urban combat action.", GenreId = 1, Genre = Genres[0] },
            new Game { Id = 7, Name = "Secrets of the Deep", Description = "An underwater adventure full of mystery.", GenreId = 2, Genre = Genres[1] },
            new Game { Id = 8, Name = "Chronicles of Aethelgard", Description = "A classic RPG with a rich storyline.", GenreId = 3, Genre = Genres[2] },
            new Game { Id = 9, Name = "Kingdoms at War", Description = "Build your kingdom and conquer your enemies.", GenreId = 4, Genre = Genres  [3] },
            new Game { Id = 10, Name = "Train Tycoon", Description = "Manage your own railway empire.", GenreId = 5, Genre = Genres[4] }
        };
    }

    public partial class Game
    {
        public int Id { get; set; }

        public string Name { get; set; } = null!;

        public string? Description { get; set; }

        public int GenreId { get; set; }

        public virtual Genre Genre { get; set; } = null!;

    }

    public partial class Genre
    {
        public int Id { get; set; }

        public string? Name { get; set; }
    }
}
