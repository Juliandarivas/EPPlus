using System;
using System.ComponentModel;

namespace EPPlus.Model.Models
{
    public class Product
    {
        public Product(int id, string referencia, string descripcion, DateTime fechaCreacion, decimal valor)
        {
            Id = id;
            Referencia = referencia;
            Descripcion = descripcion;
            FechaCreacion = fechaCreacion;
            Valor = valor;
        }

        public int Id { get; }
        public string Referencia { get; }
        [DisplayName("Descripción")]
        public string Descripcion { get; }
        [DisplayName("Fecha de creación")]
        public DateTime FechaCreacion { get; }
        public decimal Valor { get; }
        private string ReferenciaDescripcion { get; set; }
    }
}
