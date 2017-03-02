using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OnCore
{
    interface IDocx
    {
         void SetValue(string key, string value);
        bool SaveDocument(string new_document);
    }
}
