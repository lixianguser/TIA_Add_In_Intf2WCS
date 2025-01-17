﻿using Siemens.Engineering.AddIn.Menu;
using Siemens.Engineering.AddIn;
using Siemens.Engineering;
using System.Collections.Generic;

namespace TIA_Add_In_Intf2WCS
{
    public sealed class AddInProvider : ProjectTreeAddInProvider
    {
        private readonly TiaPortal _tiaPortal;

        public AddInProvider(TiaPortal tiaPortal)
        {
            _tiaPortal = tiaPortal;
        }

        protected override IEnumerable<ContextMenuAddIn> GetContextMenuAddIns()
        {
            yield return new AddIn(_tiaPortal);
        }
    }
}
