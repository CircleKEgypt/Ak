using System;
using System.Collections.Generic;

namespace CK.Models;

public partial class Ecorescategory
{
    public long Categoryhierarchy { get; set; }

    public int Changestatus { get; set; }

    public string Code { get; set; } = null!;

    public long Defaultprojectglobalcategory { get; set; }

    public decimal DefaultthresholdPsn { get; set; }

    public long Instancerelationtype { get; set; }

    public int Isactive { get; set; }

    public int Iscategoryattributesinherited { get; set; }

    public int Istangible { get; set; }

    public long Level { get; set; }

    public string Name { get; set; } = null!;

    public long Nestedsetleft { get; set; }

    public long Nestedsetright { get; set; }

    public long Parentcategory { get; set; }

    public string Pkwiucode { get; set; } = null!;

    public int? Reuseenabled { get; set; }

    public long Partition { get; set; }

    public long Recid { get; set; }

    public int Recversion { get; set; }

    public DateTime Modifieddatetime { get; set; }

    public string Modifiedby { get; set; } = null!;

    public DateTime Createddatetime { get; set; }

    public string Createdby { get; set; } = null!;

    public long Relationtype { get; set; }

    public long HsncodetableIn { get; set; }

    public long ServiceaccountingcodetableIn { get; set; }

    public int ExemptIn { get; set; }

    public int NongstIn { get; set; }

    public decimal Displayorder { get; set; }

    public string CatCustom { get; set; } = null!;

    public Guid Externalid { get; set; }
}
