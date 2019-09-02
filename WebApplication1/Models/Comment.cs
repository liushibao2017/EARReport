using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace WebApplication1.Models
{
    public class Comment
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public Guid Id { get; set; }
        public string VisitorName { get; set; }
        public string Email { get; set; }
        public DateTime Time { get; set; }
        public string Content { get; set; }

        [ForeignKey("Blog")]
        public Guid BlogId { get; set; }
        public virtual Blog Blog { get; set; }
    }
}