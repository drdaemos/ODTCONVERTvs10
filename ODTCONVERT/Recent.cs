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
        public DateTime Date { get; set; }
    }

    public static class ExtensionMethods
    {
        public static ObservableCollection<T> Remove<T>(
            this ObservableCollection<T> coll, Func<T, bool> condition)
        {
            var itemsToRemove = coll.Where(condition).ToList();

            foreach (var itemToRemove in itemsToRemove)
            {
                coll.Remove(itemToRemove);
            }

            return coll;
        }
    }

    public class RecentHolder
    {
        [UserScopedSetting()]
        [SettingsSerializeAs(System.Configuration.SettingsSerializeAs.Xml)]
        public ObservableCollection<Recent> recents { get; set; }

        public RecentHolder()
        {
            recents = new ObservableCollection<Recent>();
            configPath = Path.Combine(Environment.CurrentDirectory, "recents.xml");
            if (!File.Exists(configPath))
            {
                var file = File.Create(configPath);
                file.Close();
            }
            else Load();
        }
        private string configPath;

        private static int CompareByDate(Recent x, Recent y)
        {
            return DateTime.Compare(x.Date, y.Date)*-1;
        }

        public void Update(Recent newItem)
        {
            try
            {
                if (recents.Contains(newItem, new RecentComparer()))
                {
                    List<Recent> temp = recents.Where(item => item.Uri != newItem.Uri).ToList();
                    temp.Add(newItem);
                    temp.Sort(CompareByDate);
                    recents.Remove(x => true); 
                    temp.ForEach(x => recents.Add(x));
                }
                else
                {
                    List<Recent> temp = recents.ToList();
                    temp.Add(newItem);
                    temp.Sort(CompareByDate);
                    recents.Remove(x => true);
                    temp.ForEach(x => recents.Add(x));
                }
                Save();
            }
            catch (Exception e)
            {
                Console.Write("0");
            }
        }

        public void Save()
        {
            try
            {
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

