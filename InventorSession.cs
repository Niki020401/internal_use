using System.Runtime.InteropServices;
using Inventor;

namespace InventorStepExport
{
    /// <summary>
    /// Handles connecting to the running Inventor instance
    /// and obtaining the STEP translator add-in.
    /// </summary>
    class InventorSession
    {
        private const string StepTranslatorId =
            "{90AF7F40-0C01-11D5-8E83-0010B541CD80}";

        public Inventor.Application Application { get; private set; }

        public void Connect()
        {
            Application = (Inventor.Application)
                Marshal.GetActiveObject("Inventor.Application");
        }

        public TranslatorAddIn GetStepTranslator()
        {
            TranslatorAddIn translator = (TranslatorAddIn)
                Application.ApplicationAddIns.ItemById[StepTranslatorId];

            if (!translator.Activated)
                translator.Activate();

            return translator;
        }
    }
}
