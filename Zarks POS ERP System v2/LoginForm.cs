using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Zarks_POS_ERP_System_v2.Classes;

namespace Zarks_POS_ERP_System_v2
{
    public partial class LoginForm : Form
    {
        SetupClass setupCLS = new SetupClass();

        public LoginForm()
        {
            InitializeComponent();
            setupCLS.AuthUser();
        }

        private void LoginForm_Load(object sender, EventArgs e)
        {
            txtuserpass.Select();
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtuserpass.Text != "")
                {
                    if (txtuserpass.Text == setupCLS.admcode)
                    {
                        this.Close();

                        SetupForm setupForm = new SetupForm();
                        setupForm.Show();

                        GenerationForm gf = new GenerationForm();
                        gf.Hide();
                    }
                    else
                    {
                        MessageBox.Show("Sorry, passcode may be incorrect.", "Login Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        txtuserpass.Clear();
                        txtuserpass.Select();
                    }
                }
                else
                {
                    MessageBox.Show("Sorry, passcode may be incorrect.", "Login Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtuserpass.Clear();
                    txtuserpass.Select();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        
        private void LLX_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            GenerationForm gf = new GenerationForm();
            gf.Show();

            this.Close();
        }
    }
}
