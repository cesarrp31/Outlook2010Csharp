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

//Hecho por A.U.S. Peralta Cesar

namespace Outlook2010Csharp
{
    public class GestorOutlook
    {
        private static Outlook.Application outlook = GetApplicationObject();
        private GestorTareas gestorTareas;
        private static Outlook.Categories categorias= outlook.Session.Categories;

        private String usuario, contrasena;

        public GestorOutlook()
        {
            inicializar("","");
        }

        public GestorOutlook(String us, String cont)
        {
            inicializar(us, cont);
        }

        private void inicializar(String usuario, String contrasena)
        {
            this.usuario = usuario;
            this.contrasena = contrasena;

            //outlook = GetApplicationObject();

            /*this.gestorTareas = buscarListaTareasEnOutlook();
            Console.WriteLine("Total de tareas del usuario: " + gestorTareas.getTotalTareas());

            buscarCategoriasTareasEnOutlook();
            buscarEstadosTareasEnOutlook();*/
        }

        public void inicializarListaTareas()
        {
            this.gestorTareas = buscarListaTareasEnOutlook();
            //Console.WriteLine("Total de tareas encontradas del usuario: " + gestorTareas.getTotalTareas());
            //buscarCategoriasTareasEnOutlook();
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
                //Console.WriteLine(store.DisplayName);
                resultado.Add(store);
            }
            return resultado;
        }

        private Outlook.Store buscarStoresUsuarioEspecifico()
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
                    //Console.WriteLine("Encontrado store buscada para : "+store.DisplayName);
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
            //Console.WriteLine("Usuario a cargar tareas: "+usuario);
            List<Outlook.TaskItem> tareas;
            if (usuario == null || usuario.Equals(""))
            {
                Outlook.Folder carpetaDefaultUsuario =
                    (Outlook.Folder)outlook.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderTasks);
                tareas = buscarTareas(carpetaDefaultUsuario);
            }
            else
            {
                Outlook.Store store = buscarStoresUsuarioEspecifico();
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

        public static String getNombreOutlook()
        {
            //Console.WriteLine("aaa "+(outlook==null).ToString()+" aa "+outlook);
            return outlook.Name;
        }

        public static String getVersionOutlook()
        {
            return outlook.Version;
        }

        public String getDefaultUsuarioSesion()
        {
            return outlook.Session.CurrentUser.Name;
        }

        public static Outlook.Application getOutlook()
        {
            return outlook;
        }

        public String getUsuario()
        {
            return usuario;
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

        public static Outlook.Categories getCategorias()
        {
            return categorias;
        }

    }
    public class GestorTareas
    {
        private List<Outlook.TaskItem> tareas;

        public GestorTareas(List<Outlook.TaskItem> tareas)
        {
            tareas.Sort(
                delegate (Outlook.TaskItem x, Outlook.TaskItem y)
                    {
                        return x.CreationTime.CompareTo(y.CreationTime);
                    });
            this.tareas= tareas;
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
            //return getTarea(idx).CreationTime.ToString();
            return formatearFecha(getTarea(idx).CreationTime);
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

        private String formatearFecha(DateTime fecha)
        {
            return fecha.Day + "/" + fecha.Month + "/" + fecha.Year + " " + fecha.Hour + ":" + fecha.Minute + ":" + fecha.Second;
            //return String.Format("{0:dd/MM/yyyy HH:mm:ss tt}",fecha);
        }

    }

    
    public class GestorInterfaceOutlook
    {
        private static GestorOutlook instanciaOlkCSharp;

        private static String usuario= null, contrasena;

        private GestorInterfaceOutlook(){}

        private static GestorOutlook getInstancia()
        {
            try
            {
                TratamientoErrores.sinError();
                if (instanciaOlkCSharp == null)
                    nuevaInstanciaUsuarioOutlook();

                return instanciaOlkCSharp;
            }
            catch (SystemException e)
            {
                TratamientoErrores.registrarError(e, "Error: " + e.Message);
                return null;
            }                
        }

        private static void nuevaInstanciaUsuarioOutlook()
        {
            instanciaOlkCSharp = null;
            GestorInterfaceTareas.inicializarListaTarea();
            TratamientoErrores.sinError();
            
                if (GestorInterfaceOutlook.usuario == null || GestorInterfaceOutlook.usuario.Equals(""))
                    instanciaOlkCSharp = new GestorOutlook();
                else instanciaOlkCSharp = new GestorOutlook(GestorInterfaceOutlook.usuario, GestorInterfaceOutlook.contrasena);

                //instanciaOtkSharp.inicializarListaTareas();
        }

        public static GestorTareas getGestorTareas()
        {
            return getInstancia().getListaTareas();
        }

        [DllExport("setCredencialesUsuario", CallingConvention = CallingConvention.StdCall)]
        public static void setCredencialesUsuario(String usuario, String contrasena)
        {
            GestorInterfaceOutlook.usuario = usuario;
            GestorInterfaceOutlook.contrasena = contrasena;

            try
            {
                nuevaInstanciaUsuarioOutlook();  
            }
            catch (SystemException e)
            {
                TratamientoErrores.registrarError(e, e.Message + " " + e.Source + "\n" + e.StackTrace);
            }
        }

        [DllExport("actualizarListaTareas", CallingConvention = CallingConvention.StdCall)]
        public static void actualizarListaTareas()
        {
            try
            {
                getInstancia().inicializarListaTareas();
            }
            catch (SystemException e)
            {
                TratamientoErrores.registrarError(e, e.Message + " " + e.Source + "\n" + e.StackTrace);
            }
        }

        //[RGiesecke.DllExport.DllExport]
        [DllExport("getVersionDLL", CallingConvention = CallingConvention.StdCall)]
        public static String getVersionDLL()
        {
            return Assembly.GetExecutingAssembly().GetName().Version.ToString();
        }

        [DllExport("getNombreGestor", CallingConvention = CallingConvention.StdCall)]
        public static String getNombreSWFuente()
        {
            return GestorOutlook.getNombreOutlook().ToString();
        }

        [DllExport("getVersionGestor", CallingConvention = CallingConvention.StdCall)]
        public static String getVersionSWFuente()
        {
            return GestorOutlook.getVersionOutlook().ToString();
        }

        [DllExport("getNombreUsuario", CallingConvention = CallingConvention.StdCall)]
        public static String getNombreUsuario()
        {
            return getInstancia().getUsuario();
        }

        [DllExport("getCategoriasTarea", CallingConvention = CallingConvention.StdCall)]
        public static String getCategoriasTareaString()
        {
            StringBuilder sb = new StringBuilder();
            foreach (Outlook.Category o in GestorOutlook.getCategorias())
            {
                //Console.WriteLine(o.CategoryID+"-"+o.Name);
                sb.Append(o.Name+";");
            }
            sb.Remove(sb.Length-1, 1);
            return sb.ToString();
        }

        public static Outlook.Categories getCategoriasTarea()
        {
            return GestorOutlook.getCategorias();
        }

        [DllExport("existeError", CallingConvention = CallingConvention.StdCall)]
        public static Boolean existeError()
        {
            return TratamientoErrores.getErrorRegistrado() != null;
        }

        [DllExport("msgError", CallingConvention = CallingConvention.StdCall)]
        public static String msgError()
        {
            if (TratamientoErrores.getErrorRegistrado() == null) return "";
            else return TratamientoErrores.getErrorRegistrado().Message;
        }

        [DllExport("claseError", CallingConvention = CallingConvention.StdCall)]
        public static String claseError()
        {
            if(TratamientoErrores.getErrorRegistrado() == null) return "";
            else return TratamientoErrores.getErrorRegistrado().GetType().Name;
        }

        [DllExport("codigoError", CallingConvention = CallingConvention.StdCall)]
        public static int codigoError()
        {
            //falta desarrollar
            return TratamientoErrores.NO_EXISTE_ERROR;
        }
    }

    public class GestorInterfaceTareas
    {
        private static GestorTareas instancia;

        private GestorInterfaceTareas()
        {
            /*instancia = null;
            instancia = getInstancia()*/
        }

        public static void inicializarListaTarea()
        {
            instancia = null;
        }

        private static GestorTareas getInstancia()
        {
            try
            {
                TratamientoErrores.sinError();
                if (instancia == null)
                    instancia = GestorInterfaceOutlook.getGestorTareas();

                return instancia;
            }
            catch (SystemException e)
            {
                TratamientoErrores.registrarError(e, "Error: " + e.Message);
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
                return TratamientoErrores.tratamientoExcepcionesInt("No se puede determinar la cantidad de tareas.", ex, -1);
            }
        }

        [DllExport("getCuerpoTarea", CallingConvention = CallingConvention.StdCall)]
        public static String getCuerpoTarea(int idx)
        {
            try
            {
                String cuerpo = getInstancia().getCuerpoTarea(idx);
                return getStringUTF8(cuerpo);
            }
            catch (System.NullReferenceException ne)
            {
                Console.WriteLine("idx= "+idx+"\n"+ne.StackTrace);
                return "";
            }
            catch (SystemException ex)
            {
                return TratamientoErrores.tratamientoExcepcionesString("del Cuerpo ", ex, idx);
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
                return TratamientoErrores.tratamientoExcepcionesString("de la fecha de creación ", ex, idx);
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
                return TratamientoErrores.tratamientoExcepcionesString("de la fecha de vencimiento ", ex, idx);
            }
        }

        [DllExport("getAsuntoTarea", CallingConvention = CallingConvention.StdCall)]
        public static String getAsuntoTarea(int idx)
        {
            try
            {
                return getStringUTF8(getInstancia().getAsunto(idx));
            }
            catch (SystemException ex)
            {
                return TratamientoErrores.tratamientoExcepcionesString("del asunto ", ex, idx);
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
                return TratamientoErrores.tratamientoExcepcionesBoolean("No se puede determinar si la tarea esta completada.", ex, idx);
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
                return TratamientoErrores.tratamientoExcepcionesBoolean("No se puede determinar si la tarea esta completada.", ex, idx, false);
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
                return TratamientoErrores.tratamientoExcepcionesString("de la fecha en que la tarea fue completada ", ex, idx);
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
                return TratamientoErrores.tratamientoExcepcionesString("del propietario ", ex, idx);
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
                return TratamientoErrores.tratamientoExcepcionesString("del estado propietario ", ex, idx);
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
                return TratamientoErrores.tratamientoExcepcionesString("del delegador ", ex, idx);
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
                return TratamientoErrores.tratamientoExcepcionesBoolean("No se puede determinar si la tarea fué leída.", ex, idx, false);
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
                return TratamientoErrores.tratamientoExcepcionesBoolean("No se puede determinar si la tarea fué modificada.", ex, idx, false);
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
                return TratamientoErrores.tratamientoExcepcionesString("del estado ", ex, idx);
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
                return TratamientoErrores.tratamientoExcepcionesBoolean("No se puede guarda las modificaciones de la tarea.", ex, idx, false);
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
                return TratamientoErrores.tratamientoExcepcionesBoolean("No se puede borrar la tarea.", ex, idx, false);
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
                return TratamientoErrores.tratamientoExcepcionesInt("No se puede determinar el porcentaje de completitud de la tarea.", ex, -1);
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
                return TratamientoErrores.tratamientoExcepcionesBoolean("No se puede determinar si la tarea está en conflicto.", ex, idx, false);
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
                return TratamientoErrores.tratamientoExcepcionesBoolean("No se puede determinar si la tarea es recurrente.", ex, idx, false);
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
                return TratamientoErrores.tratamientoExcepcionesString("de la importancia ", ex, idx);
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
                return TratamientoErrores.tratamientoExcepcionesString("del id ", ex, idx);
            }
        }

        [DllExport("getCategoriaTarea", CallingConvention = CallingConvention.StdCall)]
        public static String getCategoria(int idx)
        {
            try
            {
                return getStringUTF8(getInstancia().getCategoria(idx));
            }
            catch (SystemException ex)
            {
                return TratamientoErrores.tratamientoExcepcionesString("de la categoría ", ex, idx);
            }
        }

        [DllExport("setCategoriaTarea", CallingConvention = CallingConvention.StdCall)]
        public static Boolean setCategoria(int idx, String nombre)
        {
            try
            {
                Boolean encontrado = false;
                Outlook.Category aux = null;
                foreach (Outlook.Category o in GestorInterfaceOutlook.getCategoriasTarea())
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
                return TratamientoErrores.tratamientoExcepcionesBoolean("de la categoría ", ex, idx, false);
            }
        }

        private static String getStringUTF8(String src)
        {
            //Convierte en utf8 un string
            Encoding enc = new UTF8Encoding(true, true);
            byte[] bytes = enc.GetBytes(src);
            return enc.GetString(bytes);
            //return src;
        }
    }

    class TratamientoErrores
    {
        public const int NO_EXISTE_ERROR = 0;
        public const int USUARIO_NO_ENCONTRADO = 1;
        public const int OTRO_ERROR = 10;

        private static SystemException errorRegistrado = null;
        private static DateTime fechaUltimoError = DateTime.MinValue;

        private TratamientoErrores() { }

        public static SystemException getErrorRegistrado()
        {
            return errorRegistrado;
        }

        public static void registrarError(SystemException ex, String msgConsola)
        {
            errorRegistrado = ex;
            fechaUltimoError = DateTime.Now;
            Console.Error.WriteLine(msgConsola);
        }

        public static void sinError()
        {
            errorRegistrado = null;
            fechaUltimoError = DateTime.MinValue;
        }

        public static String tratamientoExcepcionesString(String msgError, SystemException ex, int idx)
        {
            registrarError(ex, ex.GetType() + ". " + ex.Message + " " + ex.Source);
            return mensajeErrorLecturaTarea(msgError, idx);
        }

        public static Boolean tratamientoExcepcionesBoolean(String msgError, SystemException ex, int idx, Boolean respuesta)
        {
            registrarError(ex, ex.GetType() + ". " + ex.Message + " " + ex.Source + ". Valor devuelto: " + respuesta);
            return respuesta;
        }

        public static Boolean tratamientoExcepcionesBoolean(String msgError, SystemException ex, int idx)
        {
            registrarError(ex, "");
            return tratamientoExcepcionesBoolean(msgError, ex, idx, false);
        }

        public static int tratamientoExcepcionesInt(String msgError, SystemException ex, int respuesta)
        {
            registrarError(ex, ex.GetType() + ". " + ex.Message + " " + ex.Source + ". Valor devuelto: " + respuesta);
            return respuesta;
        }

        public static int tratamientoExcepcionesInt(String msgError, SystemException ex)
        {
            registrarError(ex, "");
            return tratamientoExcepcionesInt(msgError, ex, 0);
        }

        private static String mensajeErrorLecturaTarea(String parte, int nroTarea)
        {
            return "ERROR DE LECTURA " + parte + "DE LA TAREA Nº: " + nroTarea + "!!!";
        }
    }
}