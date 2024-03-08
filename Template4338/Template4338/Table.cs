using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Template4338
{
    public class MyDbContext : DbContext
    {
        public DbSet<Table> Tables { get; set; }
    }

    public class Table
    {
        public int Id { get; set; }
        public string FullName { get; set; }
        public string ClientId {get; set; }
        public string BirthDate { get; set; }
        public string Index { get; set; }
        public string City { get; set; }
        public string Street { get; set; }
        public string Home { get; set; }
        public string Apartment { get; set; }
        public string Email { get; set; }
    }
}
