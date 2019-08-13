using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.EnterpriseServices;
using System.Collections;
using System.Runtime.InteropServices;
using System.Reflection;

namespace AsteriskARI
{
   
    [Guid("AB634006-F13D-11D0-A459-004095E1DAEA")]
    internal interface IMyClass
    {
        [DispId(1)]
        //4. описываем методы которые можно будет вызывать из вне
        string Test(string mymessage);
        int TestSum(int a, int b);

    }

    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("AddIn.AsteriskARI")]
    [Description("AsteriskARI.Component")]
    public class AsteriskARI : ServicedComponent,
        IInitDone,
        ILanguageExtender,
        IMyClass
    {

        public AsteriskARI()
        {
            componentName = "AsteriskARI";
        }

        private IAsyncEvent asyncEvent = null;
        private IStatusLine statusLine = null;

        private Hashtable nameToNumber;
        private Hashtable numberToName;
        private Hashtable numberToParams;
        private Hashtable numberToRetVal;
        private Hashtable propertyNameToNumber;
        private Hashtable propertyNumberToName;
        private Hashtable numberToMethodInfoIdx;
        private Hashtable propertyNumberToPropertyInfoIdx;

        private PropertyInfo[] allPropertyInfo;
        private MethodInfo[] allMethodInfo;

        private string componentName;


        /// <summary>
        /// Инициализация компонента
        /// </summary>
        /// <param name="connection">reference to IDispatch</param>
        public void Init(
          [MarshalAs(UnmanagedType.IDispatch)]
            object connection)
        {
            asyncEvent = (IAsyncEvent)connection;
            statusLine = (IStatusLine)connection;
        }


        /// <summary>
        /// Возвращается информация о компоненте
        /// </summary>
        /// <param name="info">Component information</param>
        public void GetInfo(
          [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType=VarEnum.VT_VARIANT)]
            ref object[] info)
        {
            info[0] = 2000;
        }

        public void Done()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }



        /// <summary>
        /// Возвращается количество свойств
        /// </summary>
        /// <param name="props">Количество свойств </param>
        /// <remarks>
        /// <prototype>
        /// HRESULT GetNProps([in,out]long *plProps);
        /// </prototype>
        /// </remarks>
        public void GetNProps(ref Int32 props)
        {
            props = (Int32)propertyNameToNumber.Count;
        }

        /// <summary>
        /// Возвращает целочисленный идентификатор свойства, соответствующий 
        /// переданному имени
        /// </summary>
        /// <param name="propName">Имя свойства</param>
        /// <param name="propNum">Идентификатор свойства</param>
        /// <remarks>
        /// <prototype>
        /// HRESULT FindProp([in]BSTR bstrPropName,[in,out]long *plPropNum);
        /// </prototype>
        /// </remarks>
        public void FindProp(
          [MarshalAs(UnmanagedType.BStr)]
                String propName,
                ref Int32 propNum)
        {
            propNum = (Int32)propertyNameToNumber[propName];
        }

        /// <summary>
        /// Возвращает имя свойства, соответствующее 
        /// переданному целочисленному идентификатору
        /// </summary>
        /// <param name="propNum">Идентификатор свойства</param>
        /// <param name="propAlias"></param>
        /// <param name="propName">Имя свойства</param>
        /// <remarks>
        /// <prototype>
        /// HRESULT GetPropName([in]long lPropNum,[in]long lPropAlias,[in,out]BSTR *pbstrPropName);
        /// </prototype>
        /// </remarks>
        public void GetPropName(
          Int32 propNum,
          Int32 propAlias,
          [MarshalAs(UnmanagedType.BStr)]
            ref String propName)
        {
            propName = (String)propertyNumberToName[propNum];
        }

        public void GetPropVal(
          Int32 propNum,
          [MarshalAs(UnmanagedType.Struct)]
            ref object propVal)
        {
            propVal = allPropertyInfo[(int)propertyNumberToPropertyInfoIdx[propNum]].GetValue(this);
        }

        /// <summary>
        /// Устанавливает значение свойства.
        /// </summary>
        /// <param name="propName">Имя свойства</param>
        /// <param name="propVal">Значение свойства</param>
        /// <remarks>
        /// <prototype>
        /// HRESULT SetPropVal([in]long lPropNum,[in]VARIANT *varPropVal);
        /// </prototype>
        /// </remarks>
        public void SetPropVal(
          Int32 propNum,
          [MarshalAs(UnmanagedType.Struct)]
            ref object propVal)
        {
            allPropertyInfo[(int)propertyNumberToPropertyInfoIdx[propNum]].SetValue(this, propVal);
        }

        public void IsPropReadable(Int32 propNum, ref bool propRead)
        {
            propRead = allPropertyInfo[(int)propertyNumberToPropertyInfoIdx[propNum]].CanRead;
        }

        public void IsPropWritable(Int32 propNum, ref Boolean propWrite)
        {
            propWrite = allPropertyInfo[(int)propertyNumberToPropertyInfoIdx[propNum]].CanWrite;
        }

        /// <summary>
        /// Возвращает количество методов
        /// </summary>
        /// <param name="pMethods">Количество методов</param>
        /// <remarks>
        /// <prototype>
        /// [helpstring("method GetNMethods")]
        /// HRESULT GetNMethods([in,out]long *plMethods);
        /// </prototype>
        /// </remarks>
        public void GetNMethods(ref Int32 pMethods)
        {
            pMethods = (Int32)nameToNumber.Count;
        }

        /// <summary>
        /// Возвращает идентификатор метода
        /// </summary>
        /// <param name="methodName">Имя метода</param>
        /// <param name="methodNum">Идентификатор метода</param>
        /// <remarks>
        /// <prototype>
        /// [helpstring("method FindMethod")]
        /// HRESULT FindMethod(BSTR bstrMethodName,[in,out]long *plMethodNum);
        /// </prototype>
        /// </remarks>
        public void FindMethod(
                [MarshalAs(UnmanagedType.BStr)]
                String methodName,
                ref Int32 methodNum)
        {
            methodNum = (Int32)nameToNumber[methodName];
        }

        /// <summary>
        /// Возвращает имя метода по его идентификатору
        /// </summary>
        /// <param name="methodNum">Идентификатор метода</param>
        /// <param name="methodAlias"></param>
        /// <param name="methodName">Имя метода</param>
        /// <remarks>
        /// <prototype>
        /// HRESULT GetMethodName([in]long lMethodNum,[in]long lMethodAlias,[in,out]BSTR *pbstrMethodName);
        /// </prototype>
        /// </remarks>
        public void GetMethodName(Int32 methodNum,
          Int32 methodAlias,
          [MarshalAs(UnmanagedType.BStr)]
                ref String methodName)
        {
            methodName = (String)numberToName[methodNum];
        }


        /// <summary>
        /// Возвращает число параметров метода по его идентификатору
        /// </summary>
        /// <param name="methodNum">Идентификатор метода</param>
        /// <param name="pParams">Число параметров</param>
        /// <remarks>
        /// <prototype>
        /// HRESULT GetNParams([in]long lMethodNum,[in,out]long *plParams);
        /// </prototype>
        /// </remarks>
        public void GetNParams(Int32 methodNum, ref Int32 pParams)
        {
            pParams = (Int32)numberToParams[methodNum];
        }


        public void GetParamDefValue(
          Int32 methodNum,
          Int32 paramNum,
          [MarshalAs(UnmanagedType.Struct)]
            ref object paramDefValue)
        {
            paramDefValue = null;
        }


        /// <summary>
        /// Указывает, что у метода есть возвращаемое значение
        /// </summary>
        /// <param name="methodNum">Идентификатор метода</param>
        /// <param name="retValue">Наличие возвращаемого значения</param>
        /// <remarks>
        /// <prototype>
        /// HRESULT HasRetVal([in]long lMethodNum,[in,out]BOOL *pboolRetValue);
        /// </prototype>
        /// </remarks>
        public void HasRetVal(Int32 methodNum, ref Boolean retValue)
        {
            retValue = (Boolean)numberToRetVal[methodNum];
        }

        public void CallAsProc(
          Int32 methodNum,
          [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType=VarEnum.VT_VARIANT)]
            ref object[] pParams)
        {
            allMethodInfo[(int)numberToMethodInfoIdx[methodNum]].Invoke(
              this, pParams);
        }

        /// <summary>
        /// Вызов метода как функции с использованием идентификатора
        /// </summary>
        /// <param name="methodNum">Идентификатор метода</param>
        /// <param name="retValue">Возвращаемое значение</param>
        /// <param name="pParams">Параметры</param>
        /// <remarks>
        /// <prototype>
        /// HRESULT CallAsFunc( [in]long lMethodNum,[in,out] VARIANT *pvarRetValue,
        ///       [in] SAFEARRAY (VARIANT)*paParams);
        /// </prototype>
        /// </remarks>
        public void CallAsFunc(
        Int32 methodNum,
          [MarshalAs(UnmanagedType.Struct)]
            ref object retValue,
          [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType=VarEnum.VT_VARIANT)]
            ref object[] pParams)
        {
            retValue = allMethodInfo[(int)numberToMethodInfoIdx[methodNum]].Invoke(
              this, pParams);
        }


        /// <summary>
        /// Register component in 1C
        /// </summary>
        /// <param name="extensionName"></param>
        public void RegisterExtensionAs(
          [MarshalAs(UnmanagedType.BStr)]
                ref String extensionName)
        {
            try
            {
                // initialize data members
                nameToNumber = new Hashtable();
                numberToName = new Hashtable();
                numberToParams = new Hashtable();
                numberToRetVal = new Hashtable();

                propertyNameToNumber = new Hashtable();
                propertyNumberToName = new Hashtable();

                numberToMethodInfoIdx = new Hashtable();
                propertyNumberToPropertyInfoIdx = new Hashtable();

                // Заполнение хэш-таблиц
                Type type = this.GetType();
                Type[] allInterfaceTypes = type.GetInterfaces();

                // Определение идентификатора
                int Identifier = 0;

                foreach (Type interfaceType in allInterfaceTypes)
                {
                    if (
                      !interfaceType.Name.Equals("IDisposable")
                      && !interfaceType.Name.Equals("IManagedObject")
                      && !interfaceType.Name.Equals("IRemoteDispatch")
                      && !interfaceType.Name.Equals("IServicedComponentInfo")
                      && !interfaceType.Name.Equals("IInitDone")
                      && !interfaceType.Name.Equals("ILanguageExtender")
                      )
                    {
                        // Обработка методов
                        MethodInfo[] interfaceMethods = interfaceType.GetMethods();
                        foreach (MethodInfo interfaceMethodInfo in interfaceMethods)
                        {
                            nameToNumber.Add(interfaceMethodInfo.Name, Identifier);
                            numberToParams.Add(Identifier,
                              interfaceMethodInfo.GetParameters().Length);
                            if (typeof(void) != interfaceMethodInfo.ReturnType)
                                numberToRetVal.Add(Identifier, true);

                            Identifier++;
                        }

                        // Обработка свойств
                        PropertyInfo[] interfaceProperties = interfaceType.GetProperties();
                        foreach (PropertyInfo interfacePropertyInfo in interfaceProperties)
                        {
                            propertyNameToNumber.Add(interfacePropertyInfo.Name, Identifier);

                            Identifier++;
                        }
                    }
                }

                foreach (DictionaryEntry entry in nameToNumber)
                    numberToName.Add(entry.Value, entry.Key);
                foreach (DictionaryEntry entry in propertyNameToNumber)
                    propertyNumberToName.Add(entry.Value, entry.Key);

                // Сохранение информации о методах класса 
                allMethodInfo = type.GetMethods();

                // Сохранение информации о свойствах класса
                allPropertyInfo = type.GetProperties();

                // Отображение номера метода на индекс в массиве
                foreach (DictionaryEntry entry in numberToName)
                {
                    bool found = false;
                    for (int ii = 0; ii < allMethodInfo.Length; ii++)
                    {
                        if (allMethodInfo[ii].Name.Equals(entry.Value.ToString()))
                        {
                            numberToMethodInfoIdx.Add(entry.Key, ii);
                            found = true;
                        }
                    }
                    if (false == found)
                        throw new COMException("Метод не реализован ");
                }

                // Отображение номера свойства на индекс в массиве
                foreach (DictionaryEntry entry in propertyNumberToName)
                {
                    bool found = false;
                    for (int ii = 0; ii < allPropertyInfo.Length; ii++)
                    {
                        if (allPropertyInfo[ii].Name.Equals(entry.Value.ToString()))
                        {
                            propertyNumberToPropertyInfoIdx.Add(entry.Key, ii);
                            found = true;
                        }
                    }
                    if (false == found)
                        throw new COMException("The property is not implemented");
                }

                // Компонент инициализирован успешно
                // Возвращаем имя компонента
                extensionName = componentName;
            }
            catch (Exception)
            {
                return;
            }
        }

        public string Test(string mymessage)
        {
            return mymessage;
        }

        public int TestSum(int a, int b)
        {
            return a + b;

        }
    }


}
