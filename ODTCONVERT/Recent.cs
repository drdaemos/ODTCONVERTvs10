using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace ODTCONVERT
{
    [Serializable]
    public class Recent
    {
        public string Uri { get; set; }
        public string Name { get; set; }
    }

    public class RecentHolder
    {
        [UserScopedSetting()]
        [SettingsSerializeAs(System.Configuration.SettingsSerializeAs.Xml)]
        public ObservableCollection<Recent> recents { get; set; }

        public RecentHolder()
        {
            recents = new ObservableCollection<Recent>();
            configPath = Path.Combine(Environment.CurrentDirectory, "config.xml");
            if (!File.Exists(configPath))
            {
                var file = File.Create(configPath);
                file.Close();
            }
            else Load();
        }
        private string configPath;
        public void Save()
        {
            try {
               
            using (XmlWriter config = new XmlTextWriter(configPath, null))
            {
                System.Xml.Serialization.XmlSerializer x = new System.Xml.Serialization.XmlSerializer(recents.GetType());
                x.Serialize(config, recents);
            }
            }
            catch (Exception e)
            {
                Console.Write("0");
            }
        }
        public void Load()
        {
            try
            {
                using (XmlReader config = new XmlTextReader(configPath))
                {
                    System.Xml.Serialization.XmlSerializer x = new System.Xml.Serialization.XmlSerializer(recents.GetType());
                    ObservableCollection<Recent> temp = (ObservableCollection<Recent>)x.Deserialize(config);
                    recents = temp;
                }
            }
            catch (Exception e)
            {
                Console.Write("0");
            }
        }
    }
}

