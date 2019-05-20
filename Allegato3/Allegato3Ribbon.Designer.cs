using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Linq;

namespace Allegato3
{
    partial class Allegato3Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Variabile di progettazione necessaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Allegato3Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
            // TODO: fare una query READ sul db per tirare su FileID e Nome
            string jsonString = "{ \"templates\": { \"1\": \"Decreto legge crescita\" }}";
            JObject jsonObject = (JObject)JsonConvert.DeserializeObject(jsonString);

            for (int i = 0; i < jsonObject.First.Values().Count(); i++)
            {
                var rbTemp = this.Factory.CreateRibbonButton();
                rbTemp.Label = (string)jsonObject.First.Values().ElementAtOrDefault(i).FirstOrDefault();
                //rbTemp.Tag = query.ID; 
                rbTemp.Image = global::Allegato3.Properties.Resources.icons8_document_16;
                rbTemp.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button5_Click);
                this.gallery1.Buttons.Add((RibbonButton)rbTemp);

                var rbTemp2 = this.Factory.CreateRibbonButton();
                rbTemp2.Label = "Elimina "+ (string)jsonObject.First.Values().ElementAtOrDefault(i).FirstOrDefault();
                //rbTemp2.Tag = query.ID;
                rbTemp2.Image = global::Allegato3.Properties.Resources.icons8_delete_16;
                rbTemp2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.EliminaTemplate);
                this.gallery2.Buttons.Add((RibbonButton)rbTemp2);
            }
        }

        /// <summary> 
        /// Pulire le risorse in uso.
        /// </summary>
        /// <param name="disposing">ha valore true se le risorse gestite devono essere eliminate, false in caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Codice generato da Progettazione componenti

        /// <summary>
        /// Metodo necessario per il supporto della finestra di progettazione. Non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.gallery1 = this.Factory.CreateRibbonGallery();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.gallery2 = this.Factory.CreateRibbonGallery();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.button6 = this.Factory.CreateRibbonButton();
            this.button4 = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.button5 = this.Factory.CreateRibbonButton();
            this.editBox1 = this.Factory.CreateRibbonEditBox();
            this.button2 = this.Factory.CreateRibbonButton();
            this.group6 = this.Factory.CreateRibbonGroup();
            this.button10 = this.Factory.CreateRibbonButton();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.lblFoglioId = this.Factory.CreateRibbonLabel();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group3.SuspendLayout();
            this.group2.SuspendLayout();
            this.group5.SuspendLayout();
            this.group4.SuspendLayout();
            this.group6.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group5);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Groups.Add(this.group6);
            this.tab1.Label = "Allegato 3";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.gallery1);
            this.group1.Label = "Scarica un template";
            this.group1.Name = "group1";
            // 
            // gallery1
            // 
            this.gallery1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.gallery1.Image = global::Allegato3.Properties.Resources.icons8_bookmark_40;
            this.gallery1.Label = "Carica template";
            this.gallery1.Name = "gallery1";
            this.gallery1.ShowImage = true;
            // 
            // group3
            // 
            this.group3.Items.Add(this.gallery2);
            this.group3.Label = "Elimina il template";
            this.group3.Name = "group3";
            // 
            // gallery2
            // 
            this.gallery2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.gallery2.Image = global::Allegato3.Properties.Resources.icons8_trash_40;
            this.gallery2.Label = "Elimina template";
            this.gallery2.Name = "gallery2";
            this.gallery2.ShowImage = true;
            // 
            // group2
            // 
            this.group2.Items.Add(this.button1);
            this.group2.Items.Add(this.label1);
            this.group2.Label = "Conferma le operazioni";
            this.group2.Name = "group2";
            this.group2.Visible = false;
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = global::Allegato3.Properties.Resources.icons8_save_close_40;
            this.button1.Label = "Salva dati";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button1_Click);
            // 
            // label1
            // 
            this.label1.Label = "Salvataggio effettuato";
            this.label1.Name = "label1";
            this.label1.ShowLabel = false;
            // 
            // group5
            // 
            this.group5.Items.Add(this.button6);
            this.group5.Items.Add(this.button4);
            this.group5.Label = "Libera l\'accesso a colleghi";
            this.group5.Name = "group5";
            this.group5.Visible = false;
            // 
            // button6
            // 
            this.button6.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button6.Image = global::Allegato3.Properties.Resources.icons8_person_learning_by_reading_40;
            this.button6.Label = "Puoi soltanto leggere";
            this.button6.Name = "button6";
            this.button6.ShowImage = true;
            this.button6.Visible = false;
            // 
            // button4
            // 
            this.button4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button4.Image = global::Allegato3.Properties.Resources.icons8_edit_40;
            this.button4.Label = "Solo tu puoi modificare";
            this.button4.Name = "button4";
            this.button4.ShowImage = true;
            // 
            // group4
            // 
            this.group4.Items.Add(this.button5);
            this.group4.Items.Add(this.editBox1);
            this.group4.Items.Add(this.button2);
            this.group4.Items.Add(this.lblFoglioId);
            this.group4.Label = "Crea nuovo template";
            this.group4.Name = "group4";
            // 
            // button5
            // 
            this.button5.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button5.Image = global::Allegato3.Properties.Resources.icons8_data_sheet_filled_40__1_;
            this.button5.Label = "Salva file corrente come nuovo template";
            this.button5.Name = "button5";
            this.button5.ShowImage = true;
            this.button5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button5_Click_1);
            // 
            // editBox1
            // 
            this.editBox1.Label = "Inserisci nome del template";
            this.editBox1.Name = "editBox1";
            this.editBox1.Text = null;
            this.editBox1.Visible = false;
            this.editBox1.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.EditBox1_TextChanged);
            // 
            // button2
            // 
            this.button2.Label = "Conferma";
            this.button2.Name = "button2";
            this.button2.Visible = false;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button2_Click);
            // 
            // group6
            // 
            this.group6.Items.Add(this.button10);
            this.group6.Label = "Verifica la disponibilità";
            this.group6.Name = "group6";
            // 
            // button10
            // 
            this.button10.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button10.Image = global::Allegato3.Properties.Resources.icons8_services_40;
            this.button10.Label = "Connessione in corso";
            this.button10.Name = "button10";
            this.button10.ShowImage = true;
            // 
            // timer1
            // 
            this.timer1.Interval = 2000;
            this.timer1.Tick += new System.EventHandler(this.Timer1_Tick);
            // 
            // folderBrowserDialog1
            // 
            this.folderBrowserDialog1.HelpRequest += new System.EventHandler(this.FolderBrowserDialog1_HelpRequest);
            // 
            // lblFoglioId
            // 
            this.lblFoglioId.Label = "lblFoglioId";
            this.lblFoglioId.Name = "lblFoglioId";
            this.lblFoglioId.Visible = false;
            // 
            // Allegato3Ribbon
            // 
            this.Name = "Allegato3Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Allegato3Ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group6.ResumeLayout(false);
            this.group6.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery gallery1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button10;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
        private System.Windows.Forms.Timer timer1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button6;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery gallery2;
        internal RibbonButton button5;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        internal RibbonEditBox editBox1;
        internal RibbonButton button2;
        internal RibbonLabel lblFoglioId;
    }

    partial class ThisRibbonCollection
    {
        internal Allegato3Ribbon Allegato3Ribbon
        {
            get { return this.GetRibbon<Allegato3Ribbon>(); }
        }
    }
}
