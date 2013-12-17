using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using TSF.InteropTypes;

namespace TSF.SearchCandidateProviderInternal
{
    internal class ComReleaser : IDisposable
    {
        public T CreateComObject<T>() where T : class
        {
            var comClass = Attribute.GetCustomAttribute(typeof(T), typeof(CoClassAttribute)) as CoClassAttribute;
            if (comClass == null)
            {
                return default(T);
            }
            var obj = Activator.CreateInstance(Type.GetTypeFromCLSID(comClass.CoClass.GUID)) as T;
            RegisterObject(obj);
            return obj;
        }

        public delegate HRESULT ReceiveObjectDelegate<T>(out T obj) where T : class;

        public T ReceiveObject<T>(ReceiveObjectDelegate<T> action) where T : class
        {
            var retval = default(T);
            if (!action(out retval))
            {
                return default(T);
            }
            if (retval == null)
            {
                return default(T);
            }
            RegisterObject(retval);
            return retval;
        }

        public void RegisterObject(object o)
        {
            if (!Marshal.IsComObject(o))
            {
                return;
            }
            comObjects_.Add(o);
        }

        public void RegisterCleanup(Action action)
        {
            cleanupActions_.Add(action);
        }

        public void Dispose()
        {
            cleanupActions_.Reverse();
            cleanupActions_.ForEach(action => action());
            cleanupActions_.Clear();
            comObjects_.Reverse();
            comObjects_.ForEach(o => Marshal.ReleaseComObject(o));
            comObjects_.Clear();
        }

        private readonly List<Action> cleanupActions_ = new List<Action>();

        private readonly List<object> comObjects_ = new List<object>();
    }
}
