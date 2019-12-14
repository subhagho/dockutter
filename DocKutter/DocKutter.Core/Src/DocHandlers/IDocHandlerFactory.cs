using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocKutter.DocHandlers
{
    public interface IDocHandlerFactory
    {
        IDocHandler GetDocHandler(string name);
    }
}
