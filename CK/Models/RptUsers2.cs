using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;

namespace CK.Models;

public partial class RptUsers2
{
    public string? Storenumber { get; set; }

    public string? Username { get; set; }

    public string? Password { get; set; }

    public string? Server { get; set; }

    public string? Dmanager { get; set; }

    public string? Fmanager { get; set; }

    public string? Role { get; set; }

    public string? Department { get; set; }
    [NotMapped]
    public string? DecryptedPassword { get; set; }
}
