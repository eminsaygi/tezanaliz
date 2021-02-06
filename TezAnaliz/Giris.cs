using DevExpress.XtraEditors;
using System;
using System.Windows.Forms;

namespace TezAnaliz
{
    public partial class Giris : DevExpress.XtraEditors.XtraForm
    {
        public Giris()
        {
            InitializeComponent();
        }

        private void pictureEdit4_Click(object sender, EventArgs e)
        {
            if (textEdit1.Text.Equals("admin") && textEdit2.Text.Equals("1234"))
            {
                FrmMain frmMain = new FrmMain();
                this.Hide();

                frmMain.Show();
            }
            else
            {
                XtraMessageBox.Show("Hatalı bir kullanıcı adı veya parola girdiniz",
                "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void textEdit1_Click(object sender, EventArgs e)
        {
            textEdit1.Text = "";
        }

        private void textEdit2_Click(object sender, EventArgs e)
        {
            textEdit2.Text = "";
        }

        private void pictureEdit4_EditValueChanged(object sender, EventArgs e)
        {
        }

        private void labelControl3_Click(object sender, EventArgs e)
        {
            FrmMain frmMain = new FrmMain();
            frmMain.Show();
            this.Hide();
            XtraMessageBox.Show("Şimdilik şifresiz girmene izin verdim. Birdahakine kullanıcı adı ve şifre ile giriş yapınız.",
               "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}