using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;

namespace CK.Models;

public partial class Storeuser
{
    public string? Inventlocation { get; set; }

    public string? Storenumber { get; set; }

    public string? Username { get; set; }

    public string? Password { get; set; }

    public string? Name { get; set; }

    public string? Server { get; set; }

    public string? RmsstoNumber { get; set; }

    public int Id { get; set; }

    public string? Email { get; set; }

    public string? Dbase { get; set; }

    public string? PriceCategory { get; set; }

    public string? Franchise { get; set; }

    public string? Company { get; set; }

    public string? Zkip { get; set; }

    public string? StartDate { get; set; }

    public string? ArabicN { get; set; }

    public string? District { get; set; }

    public string? Dmanager { get; set; }

    public string? Fmanager { get; set; }
    [NotMapped]
    public string? DecryptedPassword { get; set; }
    
}
