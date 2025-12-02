using System;
using System.IO;
using System.Windows.Forms;
using DevExpress.Data.Filtering;
using DevExpress.ExpressApp;
using DevExpress.ExpressApp.Actions;
using WxyXaf.XpoExcel;

namespace WxyXaf.Demo.XpoExcelDictionary.Win.Controllers
{
    /// <summary>
    /// WinForms��ͨ��Excel���뵼��������
    /// </summary>
    public class WinExcelImportExportViewController : ExcelImportExportViewController
    {
        /// <summary>
        /// ִ�е��������ʵ��WinFormsƽ̨��Excel���빦��
        /// </summary>
        /// <param name="e">�¼�����</param>
        protected override void ExecuteImportAction(SimpleActionExecuteEventArgs e)
        {
            try
            {
                // ��ʾWinForms�ļ�ѡ��Ի���
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Title = "ѡ��Excel�ļ�";
                    openFileDialog.Filter = "Excel�ļ� (*.xlsx)|*.xlsx|Excel 97-2003�ļ� (*.xls)|*.xls|�����ļ� (*.*)|*.*";
                    openFileDialog.FilterIndex = 1;
                    openFileDialog.RestoreDirectory = true;

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string filePath = openFileDialog.FileName;

                        // ��ʾ����ģʽѡ��Ի���
                        using (var dialog = new ImportModeDialog())
                        {
                            if (dialog.ShowDialog() == DialogResult.OK)
                            {
                                ImportMode importMode = dialog.SelectedMode;

                                // ����XpoExcelHelperʵ������ע��DataDictionaryItemConverter
                            var dataDictionaryItemConverter = new WxyXaf.DataDictionaries.DataDictionaryItemConverter();
                            var excelHelper = new XpoExcelHelper(Application, null, new[] { dataDictionaryItemConverter });

                                // ִ�е���
                                var importOptions = new XpoExcelImportOptions
                                {
                                    Mode = importMode,
                                    StopOnError = false
                                };

                                // ʹ�÷�����÷��ͷ���
                                var importMethod = typeof(XpoExcelHelper).GetMethod("ImportFromExcel", new[] { typeof(string), typeof(XpoExcelImportOptions) });
                                if (importMethod == null)
                                {
                                    Application.ShowViewStrategy.ShowMessage(
                                        "�޷��ҵ�ImportFromExcel����",
                                        InformationType.Error
                                    );
                                    return;
                                }

                                var genericImportMethod = importMethod.MakeGenericMethod(ObjectType);
                                var result = (ImportResult)genericImportMethod.Invoke(excelHelper, new object[] { filePath, importOptions });

                                // ��ʾ������
                                Application.ShowViewStrategy.ShowMessage(
                                    result.HasErrors
                                        ? $"����ʧ�ܣ��ɹ�{result.SuccessCount}����ʧ��{result.FailureCount}����������Ϣ��{string.Join(Environment.NewLine, result.Errors.Select(e => e.ErrorMessage))}"
                                        : $"����ɹ�����{result.SuccessCount}����¼",
                                    result.HasErrors ? InformationType.Error : InformationType.Success
                                );

                                // ˢ����ͼ����ʾ�µ��������
                                View.RefreshDataSource();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Application.ShowViewStrategy.ShowMessage($"����Excelʧ�ܣ�{ex.Message}", InformationType.Error);
            }
        }

        /// <summary>
        /// ��д������ť����¼���ʵ��WinFormsƽ̨��Excel��������
        /// </summary>
        /// <param name="sender">�¼�������</param>
        /// <param name="e">�¼�����</param>
        protected override void ExportAction_Execute(object sender, SimpleActionExecuteEventArgs e)
        {
            try
            {
                // ʹ��XpoExcelHelper��������
                var excelHelper = new XpoExcelHelper(Application, null);
                var exportOptions = ExcelImportExportAttribute?.ExportOptions ?? new XpoExcelExportOptions();

                // ��ʾWinForms�ļ�����Ի���
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Title = "保存Excel文件";
                    saveFileDialog.Filter = "Excel文件 (*.xlsx)|*.xlsx|Excel 97-2003文件 (*.xls)|*.xls|所有文件 (*.*)|*.*";
                    saveFileDialog.FilterIndex = 1;
                    saveFileDialog.RestoreDirectory = true;
                    saveFileDialog.FileName = $"{ObjectType.Name}_导出_{DateTime.Now:yyyyMMddHHmmss}.xlsx";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string filePath = saveFileDialog.FileName;

                        // �������ݵ��ļ�
                        var exportMethod = typeof(XpoExcelHelper).GetMethod("ExportToExcel", new[] { typeof(string), typeof(CriteriaOperator), typeof(XpoExcelExportOptions) });
                        if (exportMethod == null)
                        {
                            throw new InvalidOperationException("无法找到ExportToExcel方法");
                        }

                        var genericExportMethod = exportMethod.MakeGenericMethod(ObjectType);
                        genericExportMethod.Invoke(excelHelper, new object[] { filePath, null, exportOptions });

                        // ��ʾ�ɹ���Ϣ
                        Application.ShowViewStrategy.ShowMessage(
                            $"数据已成功导出到{filePath}",
                            InformationType.Success
                        );
                    }
                }
            }
            catch (Exception ex)
            {
                // ��ʾ������Ϣ
                Application.ShowViewStrategy.ShowMessage(
                    $"����ʧ�ܣ�{ex.Message}",
                    InformationType.Error
                );
            }
        }

        /// <summary>
        /// ����ģʽѡ��Ի���
        /// </summary>
        private class ImportModeDialog : Form
        {
            private RadioButton rbCreateOnly;
            private RadioButton rbUpdateOnly;
            private RadioButton rbCreateOrUpdate;
            private RadioButton rbReplace;
            private Button btnOK;
            private Button btnCancel;
            private Label label1;

            public ImportMode SelectedMode { get; private set; }

            public ImportModeDialog()
            {
                InitializeComponent();
                SelectedMode = ImportMode.CreateAndUpdate;
                rbCreateOrUpdate.Checked = true;
            }

            private void InitializeComponent()
            {
                this.label1 = new System.Windows.Forms.Label();
                this.rbCreateOnly = new System.Windows.Forms.RadioButton();
                this.rbUpdateOnly = new System.Windows.Forms.RadioButton();
                this.rbCreateOrUpdate = new System.Windows.Forms.RadioButton();
                this.rbReplace = new System.Windows.Forms.RadioButton();
                this.btnOK = new System.Windows.Forms.Button();
                this.btnCancel = new System.Windows.Forms.Button();
                this.SuspendLayout();
                // 
                // label1
                // 
                this.label1.AutoSize = true;
                this.label1.Location = new System.Drawing.Point(12, 18);
                this.label1.Name = "label1";
                this.label1.Size = new System.Drawing.Size(82, 15);
                this.label1.TabIndex = 0;
                this.label1.Text = "��ѡ����ģʽ��";
                // 
                // rbCreateOnly
                // 
                this.rbCreateOnly.AutoSize = true;
                this.rbCreateOnly.Location = new System.Drawing.Point(30, 47);
                this.rbCreateOnly.Name = "rbCreateOnly";
                this.rbCreateOnly.Size = new System.Drawing.Size(113, 19);
                this.rbCreateOnly.TabIndex = 1;
                this.rbCreateOnly.TabStop = true;
                this.rbCreateOnly.Text = "������������";
                this.rbCreateOnly.UseVisualStyleBackColor = true;
                this.rbCreateOnly.CheckedChanged += new System.EventHandler(this.rbCreateOnly_CheckedChanged);
                // 
                // rbUpdateOnly
                // 
                this.rbUpdateOnly.AutoSize = true;
                this.rbUpdateOnly.Location = new System.Drawing.Point(30, 72);
                this.rbUpdateOnly.Name = "rbUpdateOnly";
                this.rbUpdateOnly.Size = new System.Drawing.Size(137, 19);
                this.rbUpdateOnly.TabIndex = 2;
                this.rbUpdateOnly.TabStop = true;
                this.rbUpdateOnly.Text = "�������Ѵ��ڵ�����";
                this.rbUpdateOnly.UseVisualStyleBackColor = true;
                this.rbUpdateOnly.CheckedChanged += new System.EventHandler(this.rbUpdateOnly_CheckedChanged);
                // 
                // rbCreateOrUpdate
                // 
                this.rbCreateOrUpdate.AutoSize = true;
                this.rbCreateOrUpdate.Location = new System.Drawing.Point(30, 97);
                this.rbCreateOrUpdate.Name = "rbCreateOrUpdate";
                this.rbCreateOrUpdate.Size = new System.Drawing.Size(113, 19);
                this.rbCreateOrUpdate.TabIndex = 3;
                this.rbCreateOrUpdate.TabStop = true;
                this.rbCreateOrUpdate.Text = "�������������";
                this.rbCreateOrUpdate.UseVisualStyleBackColor = true;
                this.rbCreateOrUpdate.CheckedChanged += new System.EventHandler(this.rbCreateOrUpdate_CheckedChanged);
                // 
                // rbReplace
                // 
                this.rbReplace.AutoSize = true;
                this.rbReplace.Location = new System.Drawing.Point(30, 122);
                this.rbReplace.Name = "rbReplace";
                this.rbReplace.Size = new System.Drawing.Size(113, 19);
                this.rbReplace.TabIndex = 4;
                this.rbReplace.TabStop = true;
                this.rbReplace.Text = "�滻��������";
                this.rbReplace.UseVisualStyleBackColor = true;
                this.rbReplace.CheckedChanged += new System.EventHandler(this.rbReplace_CheckedChanged);
                // 
                // btnOK
                // 
                this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
                this.btnOK.Location = new System.Drawing.Point(87, 156);
                this.btnOK.Name = "btnOK";
                this.btnOK.Size = new System.Drawing.Size(75, 23);
                this.btnOK.TabIndex = 5;
                this.btnOK.Text = "ȷ��";
                this.btnOK.UseVisualStyleBackColor = true;
                // 
                // btnCancel
                // 
                this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
                this.btnCancel.Location = new System.Drawing.Point(168, 156);
                this.btnCancel.Name = "btnCancel";
                this.btnCancel.Size = new System.Drawing.Size(75, 23);
                this.btnCancel.TabIndex = 6;
                this.btnCancel.Text = "ȡ��";
                this.btnCancel.UseVisualStyleBackColor = true;
                // 
                // ImportModeDialog
                // 
                this.AcceptButton = this.btnOK;
                this.CancelButton = this.btnCancel;
                this.ClientSize = new System.Drawing.Size(255, 191);
                this.Controls.Add(this.btnCancel);
                this.Controls.Add(this.btnOK);
                this.Controls.Add(this.rbReplace);
                this.Controls.Add(this.rbCreateOrUpdate);
                this.Controls.Add(this.rbUpdateOnly);
                this.Controls.Add(this.rbCreateOnly);
                this.Controls.Add(this.label1);
                this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
                this.MaximizeBox = false;
                this.MinimizeBox = false;
                this.Name = "ImportModeDialog";
                this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                this.Text = "ѡ����ģʽ";
                this.ResumeLayout(false);
                this.PerformLayout();
            }

            private void rbCreateOnly_CheckedChanged(object sender, EventArgs e)
            {
                if (rbCreateOnly.Checked)
                    SelectedMode = ImportMode.CreateOnly;
            }

            private void rbUpdateOnly_CheckedChanged(object sender, EventArgs e)
            {
                if (rbUpdateOnly.Checked)
                    SelectedMode = ImportMode.UpdateOnly;
            }

            private void rbCreateOrUpdate_CheckedChanged(object sender, EventArgs e)
            {
                if (rbCreateOrUpdate.Checked)
                    SelectedMode = ImportMode.CreateAndUpdate;
            }

            private void rbReplace_CheckedChanged(object sender, EventArgs e)
            {
                if (rbReplace.Checked)
                    SelectedMode = ImportMode.DeleteAndUpdate;
            }
        }
    }
}
