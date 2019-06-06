using DEWordPlugIn.Common;
using DEWordPlugIn.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;   
using System.Threading.Tasks;

namespace DEWordPlugIn.Bll
{
    public class ConfigBll
    {
        public static bool WriteConfigToTarget(Assembly assm,string proName)
        {
            string strPath = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments);
            string deWordPlugInFoder = strPath + "\\DEWordPlugIn";
            string configFile = "config.txt";
            string fullPath = Path.Combine(deWordPlugInFoder, configFile);
            string configTxt = FileHelper.GetResourceFile(assm,proName, "config.txt", "Config");
            FileHelper.CreateFolder(deWordPlugInFoder);

            if (!FileHelper.ExistsFile(fullPath))
            {
                FileHelper.CreateFile(deWordPlugInFoder, configFile);
                FileHelper.WriteFile(fullPath, configTxt);
            }
            return true;
        }

        public static bool SaveConfig(long hwnd, ConfigInfo info)
        {
            WordPlugInBll.Instance.AddDocConfig(hwnd, info);
            string strPath = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments);
            string deWordPlugInFoder = strPath + "\\DEWordPlugIn";
            string configFile = "config.txt";
            string fullPath = Path.Combine(deWordPlugInFoder, configFile);

            string json = Newtonsoft.Json.JsonConvert.SerializeObject(info);
            return FileHelper.WriteFile(fullPath, json);
        }

        /// <summary>
        /// 保存主题场景已加载读取
        /// </summary>
        /// <param name="hwnd"></param>
        /// <param name="info"></param>
        /// <returns></returns>
        public static bool SaveSceneConfig(SceneConfig info)
        {
            //WordPlugInBll.Instance.AddDocConfig(hwnd, info);
            string strPath = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments);
            string deWordPlugInFoder = strPath + "\\DEWordPlugIn";
            string configFile = $"config_{info.UserName}_{info.Url.Replace("http://", "").Replace(":",".")}.txt";
            string fullPath = Path.Combine(deWordPlugInFoder, configFile);

            string json = Newtonsoft.Json.JsonConvert.SerializeObject(info);
            return FileHelper.WriteFile(fullPath, json);
        }

        /// <summary>
        /// 主题场景已加载读取
        /// </summary>
        /// <param name="assm"></param>
        /// <param name="info"></param>
        /// <returns></returns>
        public static SceneConfig ReadSceneConfig(string  url,string userName)
        {
            string strPath = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments);
            string deWordPlugInFoder = strPath + "\\DEWordPlugIn";
            string configFile = $"config_{userName}_{url.Replace ("http://","").Replace(":", ".")}.txt";
            string fullPath = Path.Combine(deWordPlugInFoder, configFile);
            string configTxt=string.Empty ;

            if (FileHelper.ExistsFile(fullPath))
            {
                configTxt = FileHelper.ReadFile(fullPath);

            }
            SceneConfig config = Newtonsoft.Json.JsonConvert.DeserializeObject<SceneConfig>(configTxt);
            return config;
        }

            public static ConfigInfo ReadConfig(Assembly assm,string proName)
        {
            string strPath = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments);
            string deWordPlugInFoder = strPath + "\\DEWordPlugIn";
            string configFile = "config.txt";
            string fullPath = Path.Combine(deWordPlugInFoder, configFile);
            string configTxt;

            if (!FileHelper.ExistsFile(fullPath))
            {
                //不存在
                configTxt = FileHelper.GetResourceFile(assm,proName,  "config.txt", "Config");
            }
            else
            {
                configTxt = FileHelper.ReadFile(fullPath);
            }
            ConfigInfo config = Newtonsoft.Json.JsonConvert.DeserializeObject<ConfigInfo>(configTxt); 
            return config;
        }
    }
}
