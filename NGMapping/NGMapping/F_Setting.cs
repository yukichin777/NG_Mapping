using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace NGMapping
{
    public partial class F_Setting : Form
    {
        private readonly ExNumUpDown[] num_XY;
        private readonly RotateImageGenerator ImgGen;
        private readonly ExColorButton[] colorButtons = [];

        public F_Setting(List<string>NgTexts)
        {
            InitializeComponent();
            num_XY = [num_A_X1, num_A_X2, num_A_X3, num_A_X4, num_A_Y1, num_A_Y2, num_A_Y3, num_A_Y4];
            ImgGen = new RotateImageGenerator()
            {
                pBox = pictureBox1,
                RoteAngle = RotateAngle.None,
            };


            // DataGridViewの初期設定
            InitializeDataGridView();

            // ComboBoxに設定するアイテムリストを作成
            //List<string> items =
            //[
            //    "はんだボール", "はんだ屑", "炭化物",
            //    "ヒュミシール付着", "ヒュミシール未塗布",
            //    "浮き", "破損", "リードカット異常"
            //];

            colorButtons =
            [
                exColorButton1, exColorButton2, exColorButton3,
                exColorButton4, exColorButton5, exColorButton6,
                exColorButton7, exColorButton8
            ];

            // ComboBox列をDataGridViewに追加
            SetDgVCombo(NgTexts);            
        }

        private void InitializeDataGridView()
        {
            // DataGridViewの基本設定
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing; // 高さ変更禁止
            dataGridView1.ColumnHeadersHeight = 30; // ヘッダー行の高さを固定
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

        }

        private void SetDgVCombo(List<string> items)
        {
            // DataGridViewComboBoxColumnを作成
            DataGridViewComboBoxColumn comboColumn = new()
            {
                HeaderText = "NG種類",
                Name = "NG種類"
            };

            // ComboBoxのアイテムを設定
            comboColumn.Items.AddRange([.. items]);

            // DataGridViewに列を追加
            dataGridView1.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "NGワード",
                Name = "NGワード",
                Width = 180
            });
            dataGridView1.Columns.Add(comboColumn);
            dataGridView1.Columns[1].Width = 200;
        }

        private void F_Setting_Load(object sender, EventArgs e)
        {
            foreach (var item in num_XY)
            {
                item.ValueChanged += Item_ValueChanged;
                //item.ValueChanged += (s, ev) => picBox.Invalidate();
            }
            if (CSet.NgItem != null)
            {
                // CSet.NgItemをDataGridViewに復元
                foreach (var entry in CSet.NgItem)
                {
                    // NGワードとNG種類（ComboBoxのインデックス）を取得
                    string ngWord = entry.Key;
                    int selectedIndex = entry.Value;

                    // 新しい行をDataGridViewに追加
                    int rowIndex = dataGridView1.Rows.Add();
                    var row = dataGridView1.Rows[rowIndex];

                    // 各セルに値をセット
                    row.Cells["NGワード"].Value = ngWord;

                    // NG種類はコンボボックスのインデックスで指定
                    var comboCell = row.Cells["NG種類"] as DataGridViewComboBoxCell;
                    if (comboCell != null && selectedIndex >= 0 && selectedIndex < comboCell.Items.Count)
                    {
                        comboCell.Value = comboCell.Items[selectedIndex];
                    }
                }
            }

            Color[] colors = CSet.NgColors;
            for (int i = 0; i < colorButtons.Length; i++)
            {
                if (i < colors.Length)
                {
                    colorButtons[i].ButtonColor = colors[i];
                }
                else
                {
                    colorButtons[i].ButtonColor = SystemColors.Control;
                }
            }
            
            ch_DispDummyQR.Checked = CSet.FLG_DispDummyQR;

        }

        private void Item_ValueChanged(object sender, EventArgs e)
        {
            float[] xyA = [.. num_XY.Select(X => X.F_Value)];
            ImgGen.AddLine("X1", xyA[0], 0, xyA[1], 100F);
            ImgGen.AddLine("X2", xyA[2], 0, xyA[3], 100F);
            ImgGen.AddLine("Y1", 0, xyA[4], 100F, xyA[5]);
            ImgGen.AddLine("Y2", 0, xyA[6], 100F, xyA[7]);
        }

        private Rectangle GetImageDisplayRectangle(PictureBox pb)
        {
            if (pb.Image == null) return Rectangle.Empty;

            float imageRatio = (float)pb.Image.Width / pb.Image.Height;
            float controlRatio = (float)pb.Width / pb.Height;

            int width, height;
            if (imageRatio > controlRatio)
            {
                width = pb.Width;
                height = (int)(pb.Width / imageRatio);
            }
            else
            {
                height = pb.Height;
                width = (int)(pb.Height * imageRatio);
            }

            int x = (pb.Width - width) / 2;
            int y = (pb.Height - height) / 2;

            return new Rectangle(x, y, width, height);
        }
        private void Button1_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Add("新規");
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            Dictionary<string, int> ngDictionary = [];

            // DataGridViewの各行を順番に処理
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                // NGワード列の値を取得
                var textValue = row.Cells["NGワード"].Value?.ToString();

                // ComboBox列（NG種類）の値を取得
                var cell = row.Cells["NG種類"] as DataGridViewComboBoxCell;

                // ComboBoxのSelectedIndexを取得
                int? selectedIndex = cell?.Items.IndexOf(cell.Value);

                // 有効な値のみ追加
                if (!string.IsNullOrEmpty(textValue) && selectedIndex.HasValue && selectedIndex >= 0)
                {
                    ngDictionary[textValue] = selectedIndex.Value;
                }
            }

            // CSet に保存
            CSet.NgItem = ngDictionary;

            // 保存確認メッセージ
            MessageBox.Show("保存しました。");
        }

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 画像を読み込む
            if (!LoadImageToPictureBox(out string mes))
            {
                MessageBox.Show(mes);
            }
        }
        #region method----画像読み込み
        private bool LoadImageToPictureBox(out string mes)
        {
            if (comboBox1.SelectedIndex < 0) { mes = "品番未選択"; return false; }
            string Hinban = comboBox1.Text;
            string appFolderPath = Application.StartupPath;// アプリケーション自身のフォルダ
            string picFolderPath = Path.Combine(appFolderPath, "pic");
            bool selectA = radioButton1.Checked;
            float[] xyA = Hinban switch
            {
                "6365590" => selectA ? CSet.XY_6365590_A : CSet.XY_6365590_B,
                "6365630" => selectA ? CSet.XY_6365630_A : CSet.XY_6365630_B,
                _ => [30, 30, 60, 60, 30, 30, 60, 60],
            };
            for (int i = 0; i < 8; i++)
            {
                num_XY[i].Value = (decimal)xyA[i];
            }
            try
            {
                if (!Directory.Exists(picFolderPath))
                {
                    mes = $"画像ファイル保存フォルダ {picFolderPath} が存在しません。";
                    return false;
                }

                string AB = radioButton2.Checked ? "_B" : "_A";

                // 指定品番の "_A.png" と "_B.png" のファイルパスを組み立てる
                string pngFile = Path.Combine(picFolderPath, $"{Hinban}{AB}.png");

                // ファイルが存在するか確認
                if (!File.Exists(pngFile))
                {
                    mes = $"画像フォルダ内に {Hinban} のPNGファイルが存在しません。";
                    return false;
                }

                // PictureBox に画像を設定

                ImgGen.ImageFile = pngFile;
                ImgGen.AddLine("X1", xyA[0], 0, xyA[1], 100F);
                ImgGen.AddLine("X2", xyA[2], 0, xyA[3], 100F);
                ImgGen.AddLine("Y1", 0, xyA[4], 100F, xyA[5]);
                ImgGen.AddLine("Y2", 0, xyA[6], 100F, xyA[7]);


                mes = "画像の読み込み成功";
                return true;
            }
            catch (Exception ex)
            {
                // エラーの場合
                mes = $"画像の読み込み中にエラーが発生しました: {ex.Message}";
                return false;
            }
        }
        #endregion
        

        private void radioButton1_Click(object sender, EventArgs e)
        {
            // 画像を読み込む
            if (!LoadImageToPictureBox(out string mes))
            {
                MessageBox.Show(mes);
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void b_AreaSetSave_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex < 0) return;

            string Hinban = comboBox1.Text;
            bool selectA = radioButton1.Checked;

            float[] xyA = [.. num_XY.Select(x => x.F_Value)];

            switch (Hinban)
            {
                case "6365590":
                    if (selectA) { CSet.XY_6365590_A = xyA; } else { CSet.XY_6365590_B = xyA; }
                    ;
                    break;
                case "6365630":
                    if (selectA) { CSet.XY_6365630_A = xyA; } else { CSet.XY_6365630_B = xyA; }
                    break;
                default:
                    break;
            }
        }


        

        private string initFolder()
        {
            string inifol;
            if (string.IsNullOrWhiteSpace(CSet.DbPath))
            {
                inifol = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }
            else
            {
                inifol = Path.GetDirectoryName(CSet.DbPath);
                if (!Directory.Exists(inifol))
                {
                    inifol = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                }
            }
            return inifol;
        }
        private void SetDatabaseButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new()
            {
                Title = "データベースファイルを指定",
                Filter = "SQLiteデータベースファイル (*.sqlite)|*.sqlite",
                InitialDirectory = initFolder(),
            };
            if (ofd.ShowDialog() == DialogResult.OK)
            {

                SQLiteCon sqliteCon = new(ofd.FileName,CSet.DbTableCFG(), admin: false);
                if (sqliteCon.IsConnected) 
                {
                    L_DbPath.Text = ofd.FileName;
                    CSet.DbPath = ofd.FileName;
                }
                else 
                {
                    MessageBox.Show("このファイルは使うことができません。");
                }
            }
        }

        private void CreateDatabaseButton_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new()
            {
                Title = "データベースファイルを保存",
                Filter = "SQLiteデータベースファイル (*.sqlite)|*.sqlite",
                FileName = "NgMapping.sqlite",
                InitialDirectory = initFolder(),
            };

            if (saveFileDialog.ShowDialog() != DialogResult.OK)
            {
                // ユーザーがキャンセルした場合は処理を中断
                return;
            }

            // 保存先ファイルのパスを取得
            string dbPath = saveFileDialog.FileName;

            

            // SQLiteCon を使用してデータベースを作成
            try
            {
                SQLiteCon sqliteCon = new(dbPath, CSet.DbTableCFG(), admin: true);

                if (sqliteCon.IsConnected)
                {
                    L_DbPath.Text = sqliteCon.DbPath;
                    CSet.DbPath = sqliteCon.DbPath;
                }
                else
                {
                    MessageBox.Show("データベース作成に失敗しました。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"エラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        

        private void ExColorButton1_Click(object sender, EventArgs e)
        {
            ExColorButton btn = sender as ExColorButton;
            int tag = (int)btn.TagInt;

            using ColorDialog colorDialog = new();
            colorDialog.Color = btn.ButtonColor ?? SystemColors.Control;

            // ColorDialogを表示し、結果を取得する
            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                // ユーザーが色を選択した場合、その色を返す
                btn.ButtonColor = colorDialog.Color;

                Color[] colors = [..colorButtons.Select(x => x.ButtonColor??Color.Transparent)];
                CSet.NgColors = colors;
            }
        }

        private void Button7_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedCells.Count > 1 || dataGridView1.SelectedRows.Count > 1)
            {
                MessageBox.Show("削除は一行ずつしかできません。", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return; // 処理を中断
            }


            // 行が選択されている場合
            DataGridViewRow targetRow = null;

            if (dataGridView1.SelectedRows.Count == 1) // 行選択されている場合
            {
                targetRow = dataGridView1.SelectedRows[0];
            }
            else if (dataGridView1.CurrentCell != null) // カレントセルが存在する場合
            {
                targetRow = dataGridView1.CurrentCell.OwningRow;
            }

            // 削除対象の行が存在するか確認
            if (targetRow != null)
            {
                // 削除確認ダイアログを表示
                if (MessageBox.Show("選択した行を削除してもよろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    // 行を削除
                    dataGridView1.Rows.Remove(targetRow);
                }
            }
            else
            {
                // 削除対象が無い場合、メッセージを表示
                MessageBox.Show("削除する行が選択されていません。", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

        }
        private void Ch_DispDummyQR_CheckedChanged(object sender, EventArgs e)
        {
            CSet.FLG_DispDummyQR = ch_DispDummyQR.Checked;
        }
    }
}