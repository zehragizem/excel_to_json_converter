using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace exceltojson
{
    public class MyClass
    {
        // Özellikler
        // Properties
        public string Km { get; set; } = string.Empty;
        public string Istasyon { get; set; } = string.Empty;
        public string IstasyonArasiYolculuk { get; set; } = string.Empty;
        public string IstasyonBeklemeSuresi { get; set; } = string.Empty;
        public string ToplamYolculukSuresi { get; set; } = string.Empty;

        // Liste
        public List<MyList> MyList { get; set; } = new List<MyList>();

        // Yapıcı
        public MyClass() { }
    }

    public class MyList
    {
        // Özellikler
        public string Saat { get; set; } = string.Empty;
        public string ÖncekiServis { get; set; } = string.Empty;
        public string Açıklama { get; set; } = string.Empty;
        public string VarışNoktası { get; set; } = string.Empty;
        public string HatSefer { get; set; } = string.Empty;
        public string SonrakiSefer { get; set; } = string.Empty;
    

        // Yapıcı
        public MyList() { }
    }

    public class MyResults
    {
        public string Km { get; set; } = string.Empty;
        public string Istasyon { get; set; } = string.Empty;
        public string IstasyonArasiYolculuk { get; set; } = string.Empty;
        public string IstasyonBeklemeSuresi { get; set; } = string.Empty;
        public string ToplamYolculukSuresi { get; set; } = string.Empty;
        public List<MyList> MyList { get; set; } = new List<MyList>(); // Initialize the list to avoid null reference
    }

    public class ResultsCombine
    {
        public List<MyResults> MyResultsList { get; set; } = new List<MyResults>(); // Initialize the list in the declaration

        public ResultsCombine() { }
    }
}
