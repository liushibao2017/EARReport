using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace WebApplication1.Models
{
    public class Blog
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public Guid Id { get; set; }
        public string Title { get; set; }
        public string Author { get; set; }
        public DateTime Time { get; set; }
        public string Summary { get; set; }
        public string Content { get; set; }

        public virtual ICollection<Comment> Comments { get; set; } = new List<Comment>();
    }
}