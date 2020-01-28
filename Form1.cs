using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TB1300
{
    public partial class Form1 : Form
    {
        SAPbobsCOM.Company oCompany;
        string MySalesOrder;
        string MySalesInvoice;
        public Form1()
        {
            InitializeComponent();
        }


        private void btnConnect_Click(object sender, EventArgs e)
        {
            try
            {
                oCompany = new SAPbobsCOM.Company();
                oCompany.Server = "192.168.1.233:30015";
                oCompany.SLDServer = "192.168.1.233:40000";
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB;
                oCompany.CompanyDB = "SBODEMOES";
                oCompany.UserName = "manager";
                oCompany.Password = "Inicio3*";
                oCompany.DbUserName = "SYSTEM";
                oCompany.DbPassword = "Kaiser$641";

                int ret = oCompany.Connect();
                if (ret == 0)
                    MessageBox.Show("Connection ok");
                else
                    MessageBox.Show("Connection failed: " + oCompany.GetLastErrorCode().ToString() + " - " + oCompany.GetLastErrorDescription());
            }
            catch (Exception ex)
            {
                MessageBox.Show("Connection error: " + ex.Message);
            }
        }

        private void btnDisconnect_Click(object sender, EventArgs e)
        {
            try
            {
                if (oCompany.Connected == true)
                {
                    oCompany.Disconnect();
                    MessageBox.Show("You are now disconnected");
                }
                else MessageBox.Show("You are not connected to the company.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception: " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCompany);
                oCompany = null;
                Application.Exit();
            }
        }

        private void btnOrder_Click(object sender, EventArgs e)
        {
            SAPbobsCOM.Documents oSO;
            oSO = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
            try
            {              
                oSO.CardCode = "C20000";
                oSO.DocDueDate = DateTime.Today;
                oSO.Comments = "Krisztian";
                oSO.Lines.ItemCode = "A00001";
                oSO.Lines.Quantity = 2;
                oSO.Lines.Price = 100;
                oSO.Lines.Add();
                oSO.Lines.ItemCode = "A00002";
                oSO.Lines.Quantity = 1;
                oSO.Lines.Price = 50;

                int ret = oSO.Add();

                if (ret == 0)
                {
                    oCompany.GetNewObjectCode(out MySalesOrder);
                    MessageBox.Show("Add Sales Order successfull - " + MySalesOrder);
                }
                else
                {
                    MessageBox.Show("Add Sales Order failed: " + oCompany.GetLastErrorDescription());
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception: " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oSO);
                oSO = null;
            }
        }

        private void btnInvoice_Click(object sender, EventArgs e)
        {
            SAPbobsCOM.Documents oSO;
            SAPbobsCOM.Documents oInv;

            oSO = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
            oInv = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);

            try
            {
                oSO.GetByKey(Int32.Parse(MySalesOrder));
                oInv.CardCode = oSO.CardCode;

                for (int i = 0; i <= oSO.Lines.Count - 1; i++)
                {
                    oInv.Lines.BaseEntry = oSO.DocEntry;
                    oInv.Lines.BaseLine = i;
                    oInv.Lines.BaseType = 17;
                    oInv.Lines.Add();
                }

                int ret = oInv.Add();
                if (ret == 0)
                {
                    oCompany.GetNewObjectCode(out MySalesInvoice);
                    MessageBox.Show("Add Invoice successfully - " + MySalesInvoice);
                }                   
                else
                    MessageBox.Show("Add Invoice failed: " + oCompany.GetLastErrorDescription());

            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception: " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oSO);
                oSO = null;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oInv);
                oInv = null;
            }
        }

        private void btnPayment_Click(object sender, EventArgs e)
        {
            SAPbobsCOM.Payments oPay;
            oPay = (SAPbobsCOM.Payments)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
            SAPbobsCOM.Documents oInv;
            oInv = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
            try
            {
                oInv.GetByKey(Int32.Parse(MySalesInvoice));
                oPay.CardCode = oInv.CardCode;

                oPay.Invoices.DocEntry = oInv.DocEntry;
                oPay.CashAccount = "211100";
                oPay.CashSum = oInv.DocTotal;


                int ret = oPay.Add();
                if (ret == 0)
                    MessageBox.Show("Add Payment successfully");
                else
                    MessageBox.Show("Add Payment failed: " + oCompany.GetLastErrorDescription());

            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception: " + ex.Message);
            }
        }

        private void btnXml_Click(object sender, EventArgs e)
        {
            try
            {

                oCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode;
                SAPbobsCOM.Documents oInv;

                oInv = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);

                oInv.CardCode = "C20000";
                oInv.DocDueDate = DateTime.Today;
                oInv.Comments = "invoice saved as xml";
                oInv.Lines.ItemCode = "A00001";
                oInv.Lines.Quantity = 2;
                oInv.Lines.Price = 100;
                oInv.Lines.Add();
                oInv.Lines.ItemCode = "A00002";
                oInv.Lines.Quantity = 1;
                oInv.Lines.Price = 50;

                oInv.Add();
                oCompany.GetNewObjectCode(out MySalesInvoice);

                if (oInv.GetByKey(Int32.Parse(MySalesInvoice)) == true)
                {
                    oInv.SaveXML("c:\\temp\\invoice_" + MySalesInvoice + ".xml");
                    MessageBox.Show("Invoice " + MySalesInvoice + " exported to XML");
                }

                else
                    MessageBox.Show("Get Invoice failed: " + MySalesInvoice + " - " + oCompany.GetLastErrorDescription());

                string schema;
                schema = oCompany.GetBusinessObjectXmlSchema(SAPbobsCOM.BoObjectTypes.oInvoices);
                //MessageBox.Show(schema);

                oInv = oCompany.GetBusinessObjectFromXML("c:\\temp\\invoice_" + MySalesInvoice + ".xml", 0);
                oInv.Comments = "invoice loaded from xml";
                int ret = oInv.Add();

                if (ret == 0)
                {
                    oCompany.GetNewObjectCode(out MySalesInvoice);
                    MessageBox.Show("Invoice " + MySalesInvoice + " added from xml.");
                }
                    
                else
                    MessageBox.Show("Adding invoice " + MySalesInvoice + " from XML failed: " + oCompany.GetLastErrorDescription());

            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception: " + ex.Message);
            }
        }
    }
}
