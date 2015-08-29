using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;
using System.Data;
using System.IO;

namespace Word_Replacing_Tool
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        public DataTable dt_params = new DataTable();
        public DataTable dt_settings = new DataTable();
        public DataTable dt_attributes = new DataTable();
        public DataTable dt_customAttributes = new DataTable();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void readXMLs(object sender, EventArgs e)
        {
            readParam(sender, e);
            readAttributes();
            readCustomAttributes();
            readSettings();
        }

        private async void ShowMessage(string title, string text)
        {
            await this.ShowMessageAsync(title, text);
        }

        private void readParam(object sender, EventArgs e)
        {
            string paramFile = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Word Replacing Tool\parameter.xml";
            if (File.Exists(paramFile))
            {
                DataSet ds_params = new DataSet();
                ds_params.ReadXml(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Word Replacing Tool\parameter.xml");
                if(ds_params.Tables.Count == 0)
                {
                    dt_params.TableName = "Parameter";
                    dt_params.Columns.Add("Parameter Key");
                    dt_params.Columns.Add("Parameter Value");
                }
                else
                {
                    dt_params = ds_params.Tables["Parameter"];
                }
                dg_param.DataContext = dt_params.DefaultView;
            }
            else
            {
                ShowMessage("Keine Parameter gefunden", "Die Xml Datei mit den Parametern konnte nicht gefunden werden, die Standartparameter werden geladen");
                createMainXml();
            }
        }

        private void createMainXml()
        {
            dt_params.TableName = "Parameter";
            dt_params.Columns.Add("Parameter Key");
            dt_params.Columns.Add("Parameter Value");

            if (!Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Word Replacing Tool"))
            {
                Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Word Replacing Tool");
            }

            dt_params.WriteXml(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Word Replacing Tool\parameter.xml");
            dg_param.DataContext = dt_params.DefaultView;
        }

        private void readSettings()
        {
            string paramFile = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Word Replacing Tool\settings.xml";
            if (File.Exists(paramFile))
            {
                DataSet ds_settings = new DataSet();
                ds_settings.ReadXml(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Word Replacing Tool\settings.xml");
                dt_settings = ds_settings.Tables["Settings"];
                tb_OutputPath.Text = dt_settings.Rows[0][1].ToString();
                tb_OutputPattern.Text = dt_settings.Rows[1][1].ToString();
                tb_Templatepath.Text = dt_settings.Rows[2][1].ToString();
            }
            else
            {
                ShowMessage("Keine Einstellungen gefunden", "Die Xml Datei mit den Einstellungen konnte nicht gefunden werden, die Standartsettings werden geladen");
                createMainSettings();
            }
        }

        private void createMainSettings()
        {
            dt_settings.TableName = "Settings";
            dt_settings.Columns.Add("Settings Key");
            dt_settings.Columns.Add("Settings Value");

            DataRow row = dt_settings.NewRow();
            row[0] = "Outputpath";
            row[1] = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Word Replacing Tool";
            dt_settings.Rows.Add(row);

            row = dt_settings.NewRow();
            row[0] = "Outputpattern";
            row[1] = "Spezifikation_%U%_%D%_%T%_%N%";
            dt_settings.Rows.Add(row);

            row = dt_settings.NewRow();
            row[0] = "TemplatePath";
            row[1] = System.AppDomain.CurrentDomain.BaseDirectory + @"template.docx";
            dt_settings.Rows.Add(row);

            if (!Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Word Replacing Tool"))
            {
                Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Word Replacing Tool");
            }

            dt_settings.WriteXml(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Word Replacing Tool\settings.xml");
            tb_OutputPath.Text = dt_settings.Rows[0][1].ToString();
            tb_OutputPattern.Text = dt_settings.Rows[1][1].ToString();
            tb_Templatepath.Text = dt_settings.Rows[2][1].ToString();
        }

        private void readAttributes()
        {
            string attributesFile = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Word Replacing Tool\attributes.xml";
            if (File.Exists(attributesFile))
            {
                DataSet ds_attributes = new DataSet();
                ds_attributes.ReadXml(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Word Replacing Tool\attributes.xml");
                dt_attributes = ds_attributes.Tables["Attributes"];
                dg_attributes.DataContext = dt_attributes.DefaultView;
            }
            else
            {
                createMainAttributes();
            }
        }

        private void createMainAttributes()
        {
            dt_attributes.TableName = "Attributes";
            dt_attributes.Columns.Add("Attributes Key");
            dt_attributes.Columns.Add("Attributes Value");

            DataRow row = dt_attributes.NewRow();
            row[0] = "Titel";
            row[1] = String.Empty;
            dt_attributes.Rows.Add(row);

            row = dt_attributes.NewRow();
            row[0] = "Thema";
            row[1] = String.Empty;
            dt_attributes.Rows.Add(row);

            row = dt_attributes.NewRow();
            row[0] = "Autor";
            row[1] = String.Empty;
            dt_attributes.Rows.Add(row);

            row = dt_attributes.NewRow();
            row[0] = "Manager";
            row[1] = String.Empty;
            dt_attributes.Rows.Add(row);

            row = dt_attributes.NewRow();
            row[0] = "Firma";
            row[1] = String.Empty;
            dt_attributes.Rows.Add(row);

            row = dt_attributes.NewRow();
            row[0] = "Kategorie";
            row[1] = String.Empty;
            dt_attributes.Rows.Add(row);

            row = dt_attributes.NewRow();
            row[0] = "Stichwörter";
            row[1] = String.Empty;
            dt_attributes.Rows.Add(row);

            row = dt_attributes.NewRow();
            row[0] = "Kommentare";
            row[1] = String.Empty;
            dt_attributes.Rows.Add(row);

            row = dt_attributes.NewRow();
            row[0] = "Linkbasis";
            row[1] = String.Empty;
            dt_attributes.Rows.Add(row);

            if (!Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Word Replacing Tool"))
            {
                Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Word Replacing Tool");
            }

            dt_attributes.WriteXml(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Word Replacing Tool\attributes.xml");
            dg_attributes.DataContext = dt_attributes.DefaultView;
        }

        private void readCustomAttributes()
        {
            string customAttributesFile = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Word Replacing Tool\custom_attributes.xml";
            if (File.Exists(customAttributesFile))
            {
                DataSet ds_customAttributes = new DataSet();
                ds_customAttributes.ReadXml(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Word Replacing Tool\custom_attributes.xml");
                if (ds_customAttributes.Tables.Count == 0)
                {
                    dt_customAttributes.TableName = "Custom Attributes";
                    dt_customAttributes.Columns.Add("Custom Attributes Key");
                    dt_customAttributes.Columns.Add("Custom Attributes Value");
                }
                else
                {
                    dt_customAttributes = ds_customAttributes.Tables["Custom Attributes"];
                }
                dg_customAttributes.DataContext = dt_customAttributes.DefaultView;
            }
            else
            {
                ShowMessage("Keine Parameter gefunden", "Die Xml Datei mit den Parametern konnte nicht gefunden werden, die Standartparameter werden geladen");
                createMainCustomAttributes();
            }
        }

        private void createMainCustomAttributes()
        {
            dt_customAttributes.TableName = "Custom Attributes";
            dt_customAttributes.Columns.Add("Custom Attributes Key");
            dt_customAttributes.Columns.Add("Custom Attributes Value");

            if (!Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Word Replacing Tool"))
            {
                Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Word Replacing Tool");
            }

            dt_customAttributes.WriteXml(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Word Replacing Tool\custom_attributes.xml");
            dg_customAttributes.DataContext = dt_customAttributes.DefaultView;
        }

        private void saveParam(object sender, RoutedEventArgs e)
        {
            dt_params.WriteXml(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Word Replacing Tool\parameter.xml");
            ShowMessage("Gespeichert", "Die Parameter wurden gespeichert");
        }

        private void saveSettings(object sender, RoutedEventArgs e)
        {
            dt_settings.Rows[0][1] = tb_OutputPath.Text;
            dt_settings.Rows[1][1] = tb_OutputPattern.Text;
            dt_settings.Rows[2][1] = tb_Templatepath.Text;
            dt_settings.WriteXml(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Word Replacing Tool\settings.xml");
            ShowMessage("Gespeichert", "Die Einstellungen wurden gespeichert");
        }

        private void saveAttributes(object sender, RoutedEventArgs e)
        {
            dt_attributes.WriteXml(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Word Replacing Tool\attributes.xml");
            ShowMessage("Gespeichert", "Die Eigenschaften wurden gespeichert");
        }

        private void saveCustomAttributes(object sender, RoutedEventArgs e)
        {
            dt_customAttributes.WriteXml(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Word Replacing Tool\custom_attributes.xml");
            ShowMessage("Gespeichert", "Die benutzderdefinierten Eigenschaften wurden gespeichert");
        }
    }
}
