using System;

namespace EnvioCorreo_ListaSP
{
    class Program
    {
        static void Main()
        {
            EROfisis EROfisisPendiente = new EROfisis();
            try
            {
                EROfisisPendiente.EROfisisPendientes();
            }
            catch (Exception ex)
            {

                throw;
            }
        }
    }
}
