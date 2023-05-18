using XlsxMerge.Extensions;
using XlsxMerge.ViewModel;

namespace XlsxMerge.View
{
    public partial class FormWelcome : Form
    {
        private readonly DiffPathViewModel _diffPathViewModel;

        public FormWelcome(DiffPathViewModel pathViewModel)
        {
            InitializeComponent();

            _diffPathViewModel = pathViewModel;
        }

        private void textBoxPathBase_TextChanged(object sender, EventArgs e)
        {
            previewCommandLine();
        }

        private void previewCommandLine()
        {
            var resultArgs = _diffPathViewModel.Arguments(checkBoxUse3WayMerge.Checked);
            if (resultArgs == null)
                textBoxCommandline.Text = "모든 경로를 입력한 후에 예제를 확인할 수 있습니다.";
            else
                textBoxCommandline.Text = string.Join(" ", resultArgs);
        }

        private void FormWelcome_Load(object sender, EventArgs e)
        {
            this.Text = VersionName.GetFormTitleText();

            // 이벤트 핸들링을 위해 페어링
            buttonPathBase.Tag = new Action<string>(filePath => _diffPathViewModel.BasePath = filePath);
            buttonPathMine.Tag = new Action<string>(filePath => _diffPathViewModel.MinePath = filePath);
            buttonPathTheirs.Tag = new Action<string>(filePath => _diffPathViewModel.TheirsPath = filePath);
            buttonPathResult.Tag = new Action<string>(filePath => _diffPathViewModel.ResultPath = filePath);

            textBoxPathBase.BindingText(_diffPathViewModel, nameof(_diffPathViewModel.BasePath));
            textBoxPathMine.BindingText(_diffPathViewModel, nameof(_diffPathViewModel.MinePath));
            textBoxPathTheirs.BindingText(_diffPathViewModel, nameof(_diffPathViewModel.TheirsPath));
            textBoxPathResult.BindingText(_diffPathViewModel, nameof(_diffPathViewModel.ResultPath));

            previewCommandLine();
        }

        private void buttonPathXlsx_Click(object sender, EventArgs e)
        {
            Button? button = sender as Button;
            if (button == null)
                return;

            var updatePath = button.Tag as Action<string>;
            if (updatePath == null)
                return;

            if (openFileDialog1.ShowDialog() != DialogResult.OK)
                return;

            updatePath(openFileDialog1.FileName);
        }

        private void checkBoxUse3WayMerge_CheckedChanged(object sender, EventArgs e)
        {
            groupBoxTheirs.Enabled = checkBoxUse3WayMerge.Checked;
            previewCommandLine();
        }

        private void buttonStart_Click(object sender, EventArgs e)
        {
            var resultArgs = _diffPathViewModel.Arguments(checkBoxUse3WayMerge.Checked);
            if (resultArgs == null)
            {
                MessageBox.Show("모든 경로를 입력한 후에 실행이 가능합니다.");
                return;
            }
            var argumentInfo = new MergeArgumentInfo(resultArgs.ToArray());
            var formMainDiff = new FormMainDiff();
            formMainDiff.MergeArgs = argumentInfo;
            formMainDiff.ShowDialog();
            // Close();
        }
    }
}
