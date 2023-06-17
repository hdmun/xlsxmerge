using XlsxMerge.Extensions;
using XlsxMerge.Model;
using XlsxMerge.ViewModel;

namespace XlsxMerge.View
{
    public partial class FormWelcome : Form
    {
        private readonly PathViewModel _pathViewModel;

        public FormWelcome(PathViewModel pathViewModel)
        {
            InitializeComponent();

            _pathViewModel = pathViewModel;
        }

        private void textBoxPathBase_TextChanged(object sender, EventArgs e)
        {
            previewCommandLine();
        }

        private void previewCommandLine()
        {
            var resultArgs = _pathViewModel.Arguments(checkBoxUse3WayMerge.Checked);
            if (resultArgs == null)
                textBoxCommandline.Text = "모든 경로를 입력한 후에 예제를 확인할 수 있습니다.";
            else
                textBoxCommandline.Text = string.Join(" ", resultArgs);
        }

        private void FormWelcome_Load(object sender, EventArgs e)
        {
            this.Text = VersionName.GetFormTitleText();

            // 이벤트 핸들링을 위해 페어링
            buttonPathBase.Tag = new Action<string>(filePath => _pathViewModel.BasePath = filePath);
            buttonPathMine.Tag = new Action<string>(filePath => _pathViewModel.MinePath = filePath);
            buttonPathTheirs.Tag = new Action<string>(filePath => _pathViewModel.TheirsPath = filePath);
            buttonPathResult.Tag = new Action<string>(filePath => _pathViewModel.ResultPath = filePath);

            textBoxPathBase.BindingText(_pathViewModel, nameof(_pathViewModel.BasePath));
            textBoxPathMine.BindingText(_pathViewModel, nameof(_pathViewModel.MinePath));
            textBoxPathTheirs.BindingText(_pathViewModel, nameof(_pathViewModel.TheirsPath));
            textBoxPathResult.BindingText(_pathViewModel, nameof(_pathViewModel.ResultPath));
            buttonStart.BindingEnabled(_pathViewModel, nameof(_pathViewModel.EnableDiffStart));

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
            var resultArgs = _pathViewModel.Arguments(checkBoxUse3WayMerge.Checked);
            if (resultArgs == null)
            {
                MessageBox.Show("모든 경로를 입력한 후에 실행이 가능합니다.");
                return;
            }

            var argumentInfo = ProgramOptions.Parse(resultArgs.ToArray());
            _pathViewModel.DiffPathModel = DiffPathModel.From(argumentInfo);
            var diffViewModel = new DiffViewModel();
            var formMainDiff = new FormMainDiff(_pathViewModel, diffViewModel);
            formMainDiff.ShowDialog();
            // Close();
        }
    }
}
