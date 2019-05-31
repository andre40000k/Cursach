using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Cursach
{
    class Implementation
    {
        private string nameshop;
        private string obl;
        private string city;
        private string streat;
        private double number;
        private string code;
        private string nameofgoods;
        private double quantity;
        private double price;
        public int k, t;
        public string Nameshop { get => nameshop; set => nameshop = value; }
        public string Obl { get => obl; set => obl = value; }
        public string City { get => city; set => city = value; }
        public string Streat { get => streat; set => streat = value; }
        public double Number { get => number; set => number = value; }
        public string Code { get => code; set => code = value; }
        public string Nameofgoods { get => nameofgoods; set => nameofgoods = value; }
        public double Quantity { get => quantity; set => quantity = value >= 0 ? value : 0; }
        public double Price { get => price; set => price = value >= 0 ? value : 0; }
        public Implementation(){ }
        public Implementation(string nameshop, string obl, string city, string streat, double number, string code, string nameofgoods, double quantity, double price)
        {
            Nameshop = nameshop;
            Obl = obl;
            City = city;
            Streat = streat;
            Number = number;
            Code = code;
            Nameofgoods = nameofgoods;
            Quantity = quantity;
            Price = price;
        }
        public double Amount()
        {
            return quantity * price;
        }
        private void Pr1()
        {
            t = 0;
            for(int i = 0; i<nameshop.Length; ++i )
            {                
                if(nameshop[i] == '-')
                {
                    t++;
                    if (t == 2)
                    {
                        k = 1;
                        break;
                    }                        
                }
                if (nameshop[i] == '-' && Char.IsNumber(nameshop[i + 1]))
                {
                    k = 1;
                    break;
                }
            }                 
        }
        public int Proverca1()
        {
            k = 0;
            Pr1();
            if (k == 1)
                return 1;
            else
                return 0;
        }
        private void Pr2()
        {
            t = 0;
            for (int i = 0; i < obl.Length; ++i)
            {
                if (obl[i] == '-')
                {
                    t++;
                    if (t == 2)
                    {
                        k = 1;
                        break;
                    }
                }
                if (obl[i] == '-' && Char.IsNumber(obl[i + 1]))
                {
                    k = 1;
                    break;
                }
            }                    
        }
        public int Proverca2()
        {
            k = 0;
            Pr2();
            if (k != 0)
                return 1;
            else
                return 0;
        }
        private void Pr3()
        {
            t = 0;
            for (int i = 0; i < city.Length; ++i)
            {
                if (city[i] == '-')
                {
                    t++;
                    if (t == 2)
                    {
                        k = 1;
                        break;
                    }
                }
                if (city[i] == '-' && Char.IsNumber(city[i + 1]))
                {
                    k = 1;
                    break;
                }
            }                   
        }
        public int Proverca3()
        {
            k = 0;
            Pr3();
            if (k != 0)
                return 1;
            else
                return 0;
        }
        private void Pr4()
        {
            t = 0;
            for (int i = 0; i < streat.Length; ++i)
            {
                if (streat[i] == '-')
                {
                    t++;
                    if (t == 2)
                    {
                        k = 1;
                        break;
                    }
                }
                if (streat[i] == '-' && Char.IsNumber(streat[i + 1]))
                {
                    k = 1;
                    break;
                }
            }                  
        }
        public int Proverca4()
        {
            k = 0;
            Pr4();
            if (k != 0)
                return 1;
            else
                return 0;
        }
        private void Pr5()
        {
            t = 0;
            for (int i = 0; i < nameofgoods.Length; ++i)
            {
                if (nameofgoods[i] == '-')
                {
                    t++;
                    if (t == 2)
                    {
                        k = 1;
                        break;
                    }
                }
                if (nameofgoods[i] == '-' && Char.IsNumber(nameofgoods[i + 1]))
                {
                    k = 1;
                    break;
                }
            }                    
        }
        public int Proverca5()
        {
            k = 0;
            Pr5();
            if (k != 0)
                return 1;
            else
                return 0;
        }
        private void Pr6()
        {
            t = 0;
            for (int i = 0; i < code.Length; ++i)
            {
                if (code[i] == '-')
                {
                    t++;
                    if (t == 2)
                    {
                        k = 1;
                        break;
                    }
                }
                if (code[i] == '-' && Char.IsNumber(code[i + 1]))
                {
                    k = 1;
                    break;
                }
            }
        }
        public int Proverca6()
        {
            k = 0;
            Pr6();
            if (k != 0)
                return 1;
            else
                return 0;
        }
    }
}
