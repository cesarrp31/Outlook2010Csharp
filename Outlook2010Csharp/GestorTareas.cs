using RGiesecke.DllExport;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Outlook2010Csharp
{
    public class OutlookCSharp
    {
        private Outlook.Application outlook;
        private GestorTareas gestorTareas;
        private Outlook.Categories categorias;

        private String usuario, contrasena;

        public OutlookCSharp()
        {
            inicializar("","");
        }

        public OutlookCSharp(String us, String cont)
        {
            inicializar(us, cont);
        }

        private void inicializar(String usuario, String contrasena)
        {
            this.usuario = usuario;
            this.contrasena = contrasena;

            outlook = GetApplicationObject();
            this.gestorTareas = buscarListaTareasEnOutlook();

            buscarCategoriasTareasEnOutlook();
            buscarEstadosTareasEnOutlook();
        }


        private static Outlook.Application GetApplicationObject()
        {
            /*
             *Descargado de: 
             * https://msdn.microsoft.com/en-us/library/office/ff869819(v=office.15).aspx
             */
            /*
           Outlook.Application application = null;
           // Check whether there is an Outlook process running.
           if (Process.GetProcessesByName("OUTLOOK").Count() > 0)
           {
               // If so, use the GetActiveObject method to obtain the process and cast it to an Application object.
               application = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
           }
           else
           {
               // If not, create a new instance of Outlook and log on to the default profile.
               application = new Outlook.Application();
           }
           // Return the Outlook Application object.
           return application;*/

            return new Outlook.Application();
        }

        private List<Outlook.Store> buscarTodosStores()
        {
            /**
             * Permite ver todos los objetos de todos los usuarios cargados en en outlook
             * Ejemplo tomado de:
             * https://www.daniweb.com/programming/software-development/threads/475564/reading-all-store-in-outlook
             * */

            List<Outlook.Store> resultado = new List<Outlook.Store>();
            Outlook.NameSpace olNameSpace = outlook.GetNamespace("MAPI");

            foreach (Outlook.Store store in olNameSpace.Stores)
            {
                Console.WriteLine(store.DisplayName);
                resultado.Add(store);
            }
            return resultado;
        }

        private Outlook.Store buscarStoresUsuarioEspecifico(String usuario)
        {
            /**
             * Permite ver todos los objetos de todos los usuarios cargados en en outlook
             * Ejemplo tomado de:
             * https://www.daniweb.com/programming/software-development/threads/475564/reading-all-store-in-outlook
             * */

            Outlook.NameSpace olNameSpace = outlook.GetNamespace("MAPI");

            foreach (Outlook.Store store in olNameSpace.Stores)
            {
                if (store.DisplayName.StartsWith(usuario))
                {
                    Console.WriteLine(store.DisplayName);
                    return store;
                }              
            }
            throw new SystemException("Usuario no encontrado: "+usuario);
        }

        private List<Outlook.Folder> buscarTodasCarpetasStores(Outlook.Store store)
        {
            /**
             * Permite ver todos los objetos de todos los usuarios cargados en en outlook
             * Ejemplo tomado de:
             * https://www.daniweb.com/programming/software-development/threads/475564/reading-all-store-in-outlook
             * */

            Outlook.MAPIFolder rootF = store.GetRootFolder();
            //folders for store
            Outlook.Folders subF = rootF.Folders;
            List<Outlook.Folder> resultado = new List<Outlook.Folder>();
            foreach (Outlook.Folder oF in subF)
            {
                resultado.Add(oF);
            }
            return resultado;
        }

        private List<Outlook.TaskItem> buscarTareas(Outlook.Folder f)
        {
            return buscarTareas(f, new List<Outlook.TaskItem>());
        }

        private List<Outlook.TaskItem> buscarTareas(Outlook.Folder f, List<Outlook.TaskItem> resultado)
        {
            //List<Outlook.TaskItem> resultado = new List<Outlook.TaskItem>();
            foreach (object o in f.Items)
            {
                //make sure item is a mail item, not a meeting request
                if (o is Outlook.TaskItem)
                {
                    resultado.Add((Outlook.TaskItem)o);
                }
            }
            return resultado;
        }

        private GestorTareas buscarListaTareasEnOutlook()
        {
            List<Outlook.TaskItem> tareas;
            if (usuario == null || usuario.Equals(""))
            {
                Outlook.Folder carpetaDefaultUsuario =
                    (Outlook.Folder)outlook.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderTasks);
                tareas = buscarTareas(carpetaDefaultUsuario);
            }
            else
            {
                Outlook.Store store = buscarStoresUsuarioEspecifico(usuario);
                List<Outlook.Folder> carpetasUsuario = buscarTodasCarpetasStores(store);
                tareas = new List<Outlook.TaskItem>();
                foreach (Outlook.Folder f in carpetasUsuario)
                {
                    buscarTareas(f, tareas);
                }
                return new GestorTareas(tareas);
            }

            return new GestorTareas(tareas);
        }

        public String getNombreOutlook()
        {
            return this.outlook.Name;
        }

        public String getVersionOutlook()
        {
            return this.outlook.Version;
        }

        public String getDefaultUsuarioSesion()
        {
            return this.outlook.Session.CurrentUser.Name;
        }

        private void buscarCategoriasTareasEnOutlook()
        {
            categorias = outlook.Session.Categories;
        }

        private void buscarEstadosTareasEnOutlook()
        {
            //No implementado
        }

        public GestorTareas getListaTareas()
        {
            return gestorTareas;
        }

        public Outlook.Categories getCategorias()
        {
            return categorias;
        }

    }
    public class GestorTareas
    {
        private List<Outlook.TaskItem> tareas;

        public GestorTareas(List<Outlook.TaskItem> tareas)
        {
            this.tareas = tareas;            
        }

        public int getTotalTareas()
        {
            return tareas.Count;
        }

        public String getCuerpoTarea(int idx)
        {
            return getTarea(idx).Body.ToString().Trim();
        }

        public String getAsunto(int idx)
        {
            return getTarea(idx).Subject;
        }

        public String getFechaVencimiento(int idx)
        {
            return getTarea(idx).DueDate.ToString();
        }

        public String getFechaCreacion(int idx)
        {
            return getTarea(idx).CreationTime.ToString();
        }

        public Boolean isTareaCompletada(int idx)
        {
            return getTarea(idx).Complete;
        }

        public String getFechaCompletada(int idx)
        {
            return getTarea(idx).DateCompleted.ToString();
        }

        public void marcarTareaCompletada(int idx)
        {
            getTarea(idx).Complete = true;
            getTarea(idx).DateCompleted = DateTime.Now;
        }

        public String getPropietario(int idx)
        {
            return getTarea(idx).Owner.ToString();
        }

        public String getEstadoPropietario(int idx)
        {
            return getTarea(idx).Ownership.ToString();
        }

        public String getDelegador(int idx)
        {
            return getTarea(idx).Delegator.ToString();
        }

        public Boolean isLeida(int idx)
        {
            return !getTarea(idx).UnRead;
        }

        public Boolean isModificada(int idx)
        {
            return getTarea(idx).Saved;
        }

        public String getEstado(int idx)
        {
            return getTarea(idx).Status.ToString();
        }

        public void guardar(int idx)
        {
            getTarea(idx).Save();
        }

        public Boolean borrar(int idx)
        {
            return false;
        }

        public int getPorcentajeCompletada(int idx)
        {
            return getTarea(idx).PercentComplete;
        }

        public void setPorcentajeCompletada(int idx, int porc)
        {
            getTarea(idx).PercentComplete= porc;
        }

        public Boolean isConflicto(int idx)
        {
            return getTarea(idx).IsConflict;
        }

        public Boolean isRecurrente(int idx)
        {
            return getTarea(idx).IsRecurring;
        }

        public String getImportancia(int idx)
        {
            return getTarea(idx).Importance.ToString();
        }

        public String getId(int idx)
        {
            return getTarea(idx).EntryID.ToString();
        }

        public String getCategoria(int idx)
        {
            return getTarea(idx).Categories.ToString();
        }

        public void setCategoria(int idx, Outlook.Category categoria)
        {
            getTarea(idx).Categories = categoria.Name;
        }

        private Outlook.TaskItem getTarea(int idx)
        {
            if (idx >= this.getTotalTareas()) throw new IndexOutOfRangeException();
            return tareas.ElementAt(idx);
        }
    }

    
    public class SOutlook2010Sharp
    {
        public const int NO_EXISTE_ERROR = 0;
        public const int USUARIO_NO_ENCONTRADO = 1;
        public const int OTRO_ERROR = 10;

        private static OutlookCSharp instancia;

        private static String usuario= null, contrasena;

        private static SystemException errorRegistrado= null;

        private SOutlook2010Sharp(){}

        private static OutlookCSharp getInstancia()
        {
            try
            {
                if (instancia == null)
                    nuevaInstanciaUsuarioOutlook();

                return instancia;
            }
            catch (SystemException e)
            {
                registrarError(e, "Error: " + e.Message);
                //throw e;
                return null;
            }                
        }

        private static void nuevaInstanciaUsuarioOutlook()
        {
            instancia = null;
            sinError();
            try
            {
                if (SOutlook2010Sharp.usuario == null || SOutlook2010Sharp.usuario.Equals("")) instancia = new OutlookCSharp();
                else instancia = new OutlookCSharp(SOutlook2010Sharp.usuario, SOutlook2010Sharp.contrasena);
            }
            catch (SystemException e)
            {
                registrarError(e, e.Message + " " + e.Source + "\n" + e.StackTrace);
            }            
        }

        public static GestorTareas getGestorTareas()
        {
            return getInstancia().getListaTareas();
        }

        [DllExport("setCredencialesUsuario", CallingConvention = CallingConvention.StdCall)]
        public static void setUsuario(String usuario, String contrasena)
        {
            SOutlook2010Sharp.usuario = usuario;
            SOutlook2010Sharp.contrasena = contrasena;

            nuevaInstanciaUsuarioOutlook();
        }

        //[RGiesecke.DllExport.DllExport]
        [DllExport("getVersionDLL", CallingConvention = CallingConvention.StdCall)]
        public static String getVersionDLL()
        {
            return Assembly.GetExecutingAssembly().GetName().Version.ToString();
        }

        //[RGiesecke.DllExport.DllExport]
        [DllExport("getNombreGestor", CallingConvention = CallingConvention.StdCall)]
        public static String getNombreSWFuente()
        {
            return getInstancia().getNombreOutlook().ToString();
        }

        //[RGiesecke.DllExport.DllExport]
        [DllExport("getVersionGestor", CallingConvention = CallingConvention.StdCall)]
        public static String getVersionSWFuente()
        {
            return getInstancia().getVersionOutlook().ToString();
        }

        [DllExport("getNombreUsuario", CallingConvention = CallingConvention.StdCall)]
        public static String getNombreUsuario()
        {
            return getInstancia().getDefaultUsuarioSesion().ToString();
        }

        [DllExport("getCategoriasTarea", CallingConvention = CallingConvention.StdCall)]
        public static String getCategoriasTareaString()
        {
            StringBuilder sb = new StringBuilder();
            foreach (Outlook.Category o in getInstancia().getCategorias())
            {
                //Console.WriteLine(o.CategoryID+"-"+o.Name);
                sb.Append(o.Name+";");
            }
            sb.Remove(sb.Length-1, 1);
            return sb.ToString();
        }

        public static Outlook.Categories getCategoriasTarea()
        {
            return getInstancia().getCategorias();
        }

        [DllExport("existeError", CallingConvention = CallingConvention.StdCall)]
        public static Boolean existeError()
        {
            return errorRegistrado!=null;
        }

        [DllExport("msgError", CallingConvention = CallingConvention.StdCall)]
        public static String msgError()
        {
            if (errorRegistrado == null) return "";
            else return errorRegistrado.Message;
        }

        [DllExport("claseError", CallingConvention = CallingConvention.StdCall)]
        public static String claseError()
        {
            if(errorRegistrado == null) return "";
            else return errorRegistrado.GetType().Name;
        }

        [DllExport("codigoError", CallingConvention = CallingConvention.StdCall)]
        public static int codigoError()
        {
            //falta desarrolar
            return NO_EXISTE_ERROR;
        }

        public static void registrarError(SystemException ex, String msgConsola)
        {
            SOutlook2010Sharp.errorRegistrado = ex;
            Console.Error.WriteLine(msgConsola);
        }

        private static void sinError()
        {
            SOutlook2010Sharp.errorRegistrado = null;
        }

        public static String tratamientoExcepcionesString(String msgError, SystemException ex, int idx)
        {
            registrarError(ex, ex.GetType() + ". " + ex.Message + " " + ex.Source);
            return mensajeErrorLecturaTarea(msgError, idx);
        }

        public static Boolean tratamientoExcepcionesBoolean(String msgError, SystemException ex, int idx, Boolean respuesta)
        {
            registrarError(ex, ex.GetType()+". "+ex.Message + " " + ex.Source + ". Valor devuelto: "+respuesta);
            return respuesta;
        }

        public static Boolean tratamientoExcepcionesBoolean(String msgError, SystemException ex, int idx)
        {
            registrarError(ex, "");
            return tratamientoExcepcionesBoolean(msgError, ex, idx, false);
        }

        public static int tratamientoExcepcionesInt(String msgError, SystemException ex, int respuesta)
        {
            registrarError(ex, ex.GetType()+". "+ex.Message + " " + ex.Source + ". Valor devuelto: "+respuesta);
            return respuesta;
        }

        public static int tratamientoExcepcionesInt(String msgError, SystemException ex)
        {
            registrarError(ex, "");
            return tratamientoExcepcionesInt(msgError, ex, 0);
        }

        private static String mensajeErrorLecturaTarea(String parte, int nroTarea)
        {
            return "ERROR DE LECTURA "+parte+"DE LA TAREA Nº: " + nroTarea+"!!!";
        }
    }

    public class SGestorTareas
    {
        private static GestorTareas instancia;

        private SGestorTareas()
        {
            //instancia = SOutlook2010Sharp.getGestorTareas();
        }

        private static GestorTareas getInstancia()
        {
            try
            {
                if (instancia == null)
                    instancia = SOutlook2010Sharp.getGestorTareas();

                return instancia;
            }
            catch (SystemException e)
            {
                SOutlook2010Sharp.registrarError(e, "Error: " + e.Message);
                //throw e;
                return null;
            }
        }

        [DllExport("getTotalTareas", CallingConvention = CallingConvention.StdCall)]
        public static int getTotalTareas()
        {
            try
            {
                return getInstancia().getTotalTareas();
            }
            catch (SystemException ex)
            {
                return SOutlook2010Sharp.tratamientoExcepcionesInt("No se puede determinar la cantidad de tareas.", ex, -1);
            }
        }

        [DllExport("getCuerpoTarea", CallingConvention = CallingConvention.StdCall)]
        public static String getCuerpoTarea(int idx)
        {
            try
            {
                return getInstancia().getCuerpoTarea(idx);
            }
            catch (SystemException ex)
            {
                return SOutlook2010Sharp.tratamientoExcepcionesString("del Body ", ex, idx);
            }
        }

        [DllExport("getFechaCreacionTarea", CallingConvention = CallingConvention.StdCall)]
        public static String getFechaCreacionTarea(int idx)
        {
            try
            {
                return getInstancia().getFechaCreacion(idx);
            }
            catch (SystemException ex)
            {
                return SOutlook2010Sharp.tratamientoExcepcionesString("de la fecha de creación ", ex, idx);
            }
        }

        [DllExport("getFechaVencimientoTarea", CallingConvention = CallingConvention.StdCall)]
        public static String getFechaVencimientoTarea(int idx)
        {
            try
            {
                return getInstancia().getFechaVencimiento(idx);
            }
            catch (SystemException ex)
            {
                return SOutlook2010Sharp.tratamientoExcepcionesString("de la fecha de vencimiento ", ex, idx);
            }
        }

        [DllExport("getAsuntoTarea", CallingConvention = CallingConvention.StdCall)]
        public static String getAsuntoTarea(int idx)
        {
            try
            {
                return getInstancia().getAsunto(idx);
            }
            catch (SystemException ex)
            {
                return SOutlook2010Sharp.tratamientoExcepcionesString("del asunto ", ex, idx);
            }
        }

        [DllExport("isTareaCompletada", CallingConvention = CallingConvention.StdCall)]
        public static Boolean isTareaCompletada(int idx)
        {
            try
            {
                return getInstancia().isTareaCompletada(idx);
            }
            catch (SystemException ex)
            {
                return SOutlook2010Sharp.tratamientoExcepcionesBoolean("No se puede determinar si la tarea esta completada.", ex, idx);
            }
        }

        [DllExport("marcarTareaCompletada", CallingConvention = CallingConvention.StdCall)]
        public static Boolean marcarTareaCompletada(int idx)
        {
            try
            {
                getInstancia().marcarTareaCompletada(idx);
                getInstancia().setPorcentajeCompletada(idx, 100);
                getInstancia().guardar(idx);
                return true;
            }
            catch (SystemException ex)
            {
                return SOutlook2010Sharp.tratamientoExcepcionesBoolean("No se puede determinar si la tarea esta completada.", ex, idx, false);
            }
        }

        [DllExport("getFechaTareaCompletada", CallingConvention = CallingConvention.StdCall)]
        public static String getFechaCompeltada(int idx)
        {
            try
            {
                return getInstancia().getFechaCompletada(idx);
            }
            catch (SystemException ex)
            {
                return SOutlook2010Sharp.tratamientoExcepcionesString("de la fecha en que la tarea fue completada ", ex, idx);
            }
        }

        [DllExport("getPropietarioTarea", CallingConvention = CallingConvention.StdCall)]
        public static String getPropietarioTarea(int idx)
        {
            try
            {
                return getInstancia().getPropietario(idx);
            }
            catch (SystemException ex)
            {
                return SOutlook2010Sharp.tratamientoExcepcionesString("del propietario ", ex, idx);
            }
        }

        [DllExport("getEstadoPropietarioTarea", CallingConvention = CallingConvention.StdCall)]
        public static String getEstadoPropietario(int idx)
        {
            try
            {
                return getInstancia().getEstadoPropietario(idx);
            }
            catch (SystemException ex)
            {
                return SOutlook2010Sharp.tratamientoExcepcionesString("del estado propietario ", ex, idx);
            }
        }

        [DllExport("getDelegadorTarea", CallingConvention = CallingConvention.StdCall)]
        public static String getDelegador(int idx)
        {
            try
            {
                return getInstancia().getDelegador(idx);
            }
            catch (SystemException ex)
            {
                return SOutlook2010Sharp.tratamientoExcepcionesString("del delegador ", ex, idx);
            }
        }

        [DllExport("isTareaLeida", CallingConvention = CallingConvention.StdCall)]
        public static Boolean isLeida(int idx)
        {
            try
            {
                return getInstancia().isLeida(idx); ;
            }
            catch (SystemException ex)
            {
                return SOutlook2010Sharp.tratamientoExcepcionesBoolean("No se puede determinar si la tarea fué leída.", ex, idx, false);
            }
        }

        [DllExport("isTareaModificada", CallingConvention = CallingConvention.StdCall)]
        public static Boolean isModificada(int idx)
        {
            try
            {
                return getInstancia().isModificada(idx); ;
            }
            catch (SystemException ex)
            {
                return SOutlook2010Sharp.tratamientoExcepcionesBoolean("No se puede determinar si la tarea fué modificada.", ex, idx, false);
            }
        }

        [DllExport("getEstadoTarea", CallingConvention = CallingConvention.StdCall)]
        public static String getEstado(int idx)
        {
            try
            {
                return getInstancia().getEstado(idx);
            }
            catch (SystemException ex)
            {
                return SOutlook2010Sharp.tratamientoExcepcionesString("del estado ", ex, idx);
            }
        }

        [DllExport("guardarModificacionTarea", CallingConvention = CallingConvention.StdCall)]
        public static Boolean guardar(int idx)
        {
            try
            {
                getInstancia().guardar(idx);
                return true;
            }
            catch (SystemException ex)
            {
                return SOutlook2010Sharp.tratamientoExcepcionesBoolean("No se puede guarda las modificaciones de la tarea.", ex, idx, false);
            }
        }

        [DllExport("borrarTarea", CallingConvention = CallingConvention.StdCall)]
        public static Boolean borrar(int idx)
        {
            try
            {
                return getInstancia().borrar(idx); ;
            }
            catch (SystemException ex)
            {
                return SOutlook2010Sharp.tratamientoExcepcionesBoolean("No se puede borrar la tarea.", ex, idx, false);
            }
        }

        [DllExport("getPorcentajeTareaCompletada", CallingConvention = CallingConvention.StdCall)]
        public static int getPorcentajeCompletada(int idx)
        {
            try
            {
                return getInstancia().getPorcentajeCompletada(idx);
            }
            catch (SystemException ex)
            {
                return SOutlook2010Sharp.tratamientoExcepcionesInt("No se puede determinar el porcentaje de completitud de la tarea.", ex, -1);
            }
        }

        [DllExport("isTareaConflicto", CallingConvention = CallingConvention.StdCall)]
        public static Boolean isConflicto(int idx)
        {
            try
            {
                return getInstancia().isConflicto(idx); ;
            }
            catch (SystemException ex)
            {
                return SOutlook2010Sharp.tratamientoExcepcionesBoolean("No se puede determinar si la tarea está en conflicto.", ex, idx, false);
            }
        }

        [DllExport("isTareaRecurrente", CallingConvention = CallingConvention.StdCall)]
        public static Boolean isRecurrente(int idx)
        {
            try
            {
                return getInstancia().isRecurrente(idx); ;
            }
            catch (SystemException ex)
            {
                return SOutlook2010Sharp.tratamientoExcepcionesBoolean("No se puede determinar si la tarea es recurrente.", ex, idx, false);
            }
        }

        [DllExport("getImportanciaTarea", CallingConvention = CallingConvention.StdCall)]
        public static String getImportancia(int idx)
        {
            try
            {
                return getInstancia().getImportancia(idx);
            }
            catch (SystemException ex)
            {
                return SOutlook2010Sharp.tratamientoExcepcionesString("de la importancia ", ex, idx);
            }
        }

        [DllExport("getIdTarea", CallingConvention = CallingConvention.StdCall)]
        public static String getId(int idx)
        {
            try
            {
                return getInstancia().getId(idx);
            }
            catch (SystemException ex)
            {
                return SOutlook2010Sharp.tratamientoExcepcionesString("del id ", ex, idx);
            }
        }

        [DllExport("getCategoriaTarea", CallingConvention = CallingConvention.StdCall)]
        public static String getCategoria(int idx)
        {
            try
            {
                return getInstancia().getCategoria(idx);
            }
            catch (SystemException ex)
            {
                return SOutlook2010Sharp.tratamientoExcepcionesString("de la categoría ", ex, idx);
            }
        }

        [DllExport("setCategoriaTarea", CallingConvention = CallingConvention.StdCall)]
        public static Boolean setCategoria(int idx, String nombre)
        {
            try
            {
                Boolean encontrado = false;
                Outlook.Category aux = null;
                foreach (Outlook.Category o in SOutlook2010Sharp.getCategoriasTarea())
                {
                    //Console.WriteLine(o.Name+" "+nombre);
                    if (o.Name.Equals(nombre))
                    {
                        aux = o;
                        encontrado = true;
                        break;
                    }
                }
                if (encontrado) getInstancia().setCategoria(idx, aux);
                //getInstancia().getListaTareas().guardar(idx);
                return encontrado;
            }
            catch (SystemException ex)
            {
                return SOutlook2010Sharp.tratamientoExcepcionesBoolean("de la categoría ", ex, idx, false);
            }
        }
    }
}