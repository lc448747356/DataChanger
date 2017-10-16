using UnityEngine;
using System.Collections;
using System.Reflection;
using System.Collections.Generic;
public class DataManager : Singeleton<DataManager>
{
    /****************************************/
    /*使用前添加同名下划线词典*/
    /****************************************/
  
    #region Fields	
    //Resources文件夹下的存放Data目录的名称
    const string filePath = "Data";
    public Dictionary<int,DataTest> _DataTest;
    #endregion
    #region Functions	
    private void LoadText()
    {
        Object[] textObjs = Resources.LoadAll(filePath, typeof(TextAsset));
        MethodInfo normalMethod = typeof(DataManager).GetMethod("MyDataToObj",BindingFlags.Instance|BindingFlags.NonPublic);
        for (int i = 0; i < textObjs.Length; i++)
        {
            TextAsset dataText = textObjs[i] as TextAsset;
            System.Type instanceType = System.Type.GetType(dataText.name);
            MethodInfo genericMethod = normalMethod.MakeGenericMethod(instanceType);
            genericMethod.Invoke(this, new object[] { dataText});
        }
    }
    private void MyDataToObj<T>(TextAsset text)
    {
        string typeName = text.name;
        string[] lineStr = text.ToString().Split(new string[] { "\r\n" }, System.StringSplitOptions.RemoveEmptyEntries);
        string[,] gridStr = new string[lineStr.Length, lineStr[0].Split('^').Length];
        for (int i = 0; i < gridStr.GetLength(0); i++)
        {
            string[] tempGridStr = lineStr[i].Split('^');
            for (int j = 0; j < gridStr.GetLength(1); j++)
            {
                gridStr[i, j] = tempGridStr[j];
            }
        }
        System.Type instanceType = System.Type.GetType(typeName);
        FieldInfo[] fileds = instanceType.GetFields();
        //确定grid字段索引对应
        Dictionary<int, int> index = new Dictionary<int, int>();
        Dictionary<int, string> types = new Dictionary<int, string>();
        for (int i = 1; i < gridStr.GetLength(1); i++)
        {
            for (int j = 0; j < fileds.Length; j++)
            {
                if (gridStr[0, i].Equals(fileds[j].Name))
                {
                    index.Add(i, j);
                    types.Add(i, fileds[j].FieldType.ToString());
                    break;
                }
            }
        }
        Dictionary<int, T> tempDic = new Dictionary<int, T>();
        //字段赋值
        for (int i = 1; i < gridStr.GetLength(0); i++)
        {
            T instance = System.Activator.CreateInstance<T>();
            for (int j = 1; j < gridStr.GetLength(1); j++)
            {
                switch (types[j])
                {
                    case "System.Int32":
                        fileds[index[j]].SetValue(instance, int.Parse(gridStr[i, j]));
                        break;
                    case "System.String":
                        fileds[index[j]].SetValue(instance, gridStr[i, j]);
                        break;
                    case "System.Single":
                        fileds[index[j]].SetValue(instance, float.Parse(gridStr[i, j]));
                        break;
                    case "System.Single[]":
                        string[] _StrArray = gridStr[i, j].Split(';');
                        float[] tempFloatArray = new float[_StrArray.Length];
                        for (int k = 0; k < tempFloatArray.Length; k++)
                        {
                            tempFloatArray[k] = float.Parse(_StrArray[k]);
                        }
                        fileds[index[j]].SetValue(instance, tempFloatArray);
                        break;
                    case "System.Int32[]":
                        string[] __StrArray = gridStr[i, j].Split(';');
                        int[] tempIntArray = new int[__StrArray.Length];
                        for (int k = 0; k < tempIntArray.Length; k++)
                        {
                            tempIntArray[k] = int.Parse(__StrArray[k]);
                        }
                        fileds[index[j]].SetValue(instance, tempIntArray);
                        break;
                    default:
                        break;
                }
            }
            tempDic.Add(int.Parse(gridStr[i, 0]), instance);
        }
        typeof(DataManager).GetField("_" + typeName).SetValue(this,tempDic);
    }
    protected override void InitiateData()
    {
        LoadText();
    }
    #endregion



}
