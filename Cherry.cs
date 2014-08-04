using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using ComIDataObject = System.Runtime.InteropServices.ComTypes.IDataObject;

namespace Cherry
{
    public partial class Cherry : Form
    {
        /*** Variables ***/

        public const int c_iOpenedWidth = 300;
        public const int c_iOpenedHeight = 651;
        public const int c_iClosedWidth = 150;
        public const int c_iClosedHeight = c_iClosedWidth - 4;
        public const int c_iCherryBorder = 19;

        private bool m_FirstActivation;
        private bool m_bOpened;
        private string m_sClipboard;
        private ResourceInfo m_resInfo = new ResourceInfo();


        /*** Methods ***/

        public Cherry()
        {
            // Default size: 300; 296
            InitializeComponent();

            // Initialize State
            this.m_FirstActivation = true;
            this.menuStrip.Visible = false;
            this.Opened = false;
        }

        public void CherryOpen(string a_sLocation)
        {
            m_resInfo.Location = a_sLocation;
            dataGridView.Rows.Clear();
            foreach (KeyValuePair<string, string> item in m_resInfo.Resource)
            {
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(this.dataGridView);
                row.Cells[0].Value = "Add"; // this makes it work
                this.dataGridView.Rows.Add(row); // #3 need this before adding data
                row.Cells["Property"].Value = item.Key;
                row.Cells["Value"].Value = item.Value;
                row.Dispose();
            }
            this.Opened = true;
        }

        public void CherryClose()
        {
            dataGridView.Rows.Clear();
            this.Opened = false;
        }

        private void PrepareDropEvent(ref DragEventArgs e, ref Win32Point wp)
        {
            // Only allow file drop
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
            else
                e.Effect = DragDropEffects.None;

            Point p = Cursor.Position;
            wp.x = p.X;
            wp.y = p.Y;
        }


        /*** ToolStripMenu Events ***/

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                CherryOpen(this.openFileDialog.FileName);
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                using (StreamWriter outfile = new StreamWriter(this.saveFileDialog.FileName))
                {
                    outfile.WriteLine("Property;Value");
                    foreach (KeyValuePair<string, string> item in m_resInfo.Resource)
                        outfile.WriteLine(item.Key + ";" + item.Value);
                }
            }
        }

        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CherryClose();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            CherryOpen(assembly.Location);
        }


        /*** Form Events ***/

        private void Cherry_Activated(object sender, EventArgs e)
        {
            if (m_FirstActivation)
            {
                m_FirstActivation = false;
                return;
            }

            if (Clipboard.ContainsText())
            {
                string sLocation = Clipboard.GetText();
                if (sLocation != m_sClipboard)
                {
                    m_sClipboard = sLocation;
                    if (File.Exists(sLocation) &&
                        MessageBox.Show("Import this file from clipboard?\n" + sLocation, "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        CherryOpen(sLocation);
                }
            }
        }

        private void Cherry_SizeChanged(object sender, EventArgs e)
        {
            this.dataGridView.Columns[1].Width = this.Width - this.dataGridView.Columns[0].Width - c_iCherryBorder;
        }

        private void dataGridView_ColumnWidthChanged(object sender, DataGridViewColumnEventArgs e)
        {
            if (e.Column.Name == "Property")
                Cherry_SizeChanged(sender, e);
        }


        /*** Drag Events ***/

        private void Cherry_DragEnter(object sender, DragEventArgs e)
        {
            Win32Point wp;
            wp.x = wp.y = 0;
            PrepareDropEvent(ref e, ref wp);
            IDropTargetHelper dropHelper = (IDropTargetHelper)new DragDropHelper();
            try
            {
                dropHelper.DragEnter(this.Handle, (ComIDataObject)e.Data, ref wp, (int)e.Effect);
            }
            catch (COMException exception)
            {
                Console.WriteLine("Drag source is not a ComIDataObject.");
            }
        }

        private void Cherry_DragOver(object sender, DragEventArgs e)
        {
            Win32Point wp;
            wp.x = wp.y = 0;
            PrepareDropEvent(ref e, ref wp);
            IDropTargetHelper dropHelper = (IDropTargetHelper)new DragDropHelper();
            dropHelper.DragOver(ref wp, (int)e.Effect);
        }

        private void Cherry_DragLeave(object sender, EventArgs e)
        {
            IDropTargetHelper dropHelper = (IDropTargetHelper)new DragDropHelper();
            dropHelper.DragLeave();
        }

        private void Cherry_DragDrop(object sender, DragEventArgs e)
        {
            Win32Point wp;
            wp.x = wp.y = 0;
            PrepareDropEvent(ref e, ref wp);
            IDropTargetHelper dropHelper = (IDropTargetHelper)new DragDropHelper();
            dropHelper.Drop((ComIDataObject)e.Data, ref wp, (int)e.Effect);

            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            foreach (string file in files)
            {
                if (File.Exists(file))
                {
                    CherryOpen(file);
                    this.Activate();
                    break;
                }
            }
        }


        /**** Property ***/

        private bool Opened
        {
            get
            {
                return m_bOpened;
            }
            set
            {
                m_bOpened = value;
                this.saveToolStripMenuItem.Enabled = value;
                this.closeToolStripMenuItem.Enabled = value;
                this.dataGridView.Visible = value;
                this.TopMost = !value;
                if (m_bOpened)
                {
                    this.Width = c_iOpenedWidth;
                    this.Height = c_iOpenedHeight;
                }
                else
                {
                    this.Width = c_iClosedWidth;
                    this.Height = c_iClosedHeight;
                }
                if (this.menuStrip.Visible)
                    this.Height += this.menuStrip.Height;
            }
        }
    }
}
