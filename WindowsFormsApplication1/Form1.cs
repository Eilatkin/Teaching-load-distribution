using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApplication1
{

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        // =======================================================================
        // Реализация Венгерского алгоритма применительно к входной матрице matrix
        // =======================================================================

        public sealed class HungarianAlgorithm
	{
		private readonly int[,] _costMatrix;
		private int _inf;
		private int _n; //размер таблицы
		private int[] _lx; //метки преподавателей
		private int[] _ly; //метки предметов
		private bool[] _s;
		private bool[] _t;
		private int[] _matchX; //вершина соответствующая x
		private int[] _matchY; //вершина соответсвующая y
		private int _maxMatch;
		private int[] _slack;
		private int[] _slackx;
		private int[] _prev; //запоминаемые пути


		public HungarianAlgorithm(int[,] costMatrix)
		{
			_costMatrix = costMatrix;
		}

		
		public int[] Run()
		{
			_n = _costMatrix.GetLength(0);

			_lx = new int[_n];
			_ly = new int[_n];
			_s = new bool[_n];
			_t = new bool[_n];
			_matchX = new int[_n];
			_matchY = new int[_n];
			_slack = new int[_n];
			_slackx = new int[_n];
			_prev = new int[_n];
			_inf = int.MaxValue;


			InitMatches();

			if (_n != _costMatrix.GetLength(1))
				return null;

			InitLbls();

			_maxMatch = 0;

			InitialMatching();

			var q = new Queue<int>();

			#region augment

			while (_maxMatch != _n)
			{
				q.Clear();

				InitSt();

				//запоминаем позицию корневой вершины и вдух других
				var root = 0;
				int x;
				var y = 0;

				//выбирает корень из трёх вершин
				for (x = 0; x < _n; x++)
				{
					if (_matchX[x] != -1) continue;
					q.Enqueue(x);
					root = x;
					_prev[x] = -2;

					_s[x] = true;
					break;
				}

				//init slack
				for (var i = 0; i < _n; i++)
				{
					_slack[i] = _costMatrix[root, i] - _lx[root] - _ly[i];
					_slackx[i] = root;
				}

				//нахождение дополняющего пути
				while (true)
				{
					while (q.Count != 0)
					{
						x = q.Dequeue();
						var lxx = _lx[x];
						for (y = 0; y < _n; y++)
						{
							if (_costMatrix[x, y] != lxx + _ly[y] || _t[y]) continue;
							if (_matchY[y] == -1) break; //дополняющий путь найден
							_t[y] = true;
							q.Enqueue(_matchY[y]);

							AddToTree(_matchY[y], x);
						}
						if (y < _n) break; //дополняющий путь найден
					}
					if (y < _n) break; //дополняющий путь найден
					UpdateLabels(); //дополняющий путь не найден, обновляем метки

					for (y = 0; y < _n; y++)
					{
						//в этом цикле мы добавляем ребра которые были добавлены в граф равенств 
						//в результате улучшения метокre. мы добавляем ребро (slackx[y], y) к дереву 
						//в том и только том случае если !T[y] &&  slack[y] == 0, 
						//кроме того, добвляем ещё ребро
						//(y, yx[y]) либо дополняем подходящее, если y был открытым

						if (_t[y] || _slack[y] != 0) continue;
						if (_matchY[y] == -1) //найдена открытай вершина - существует дополняющий путь
						{
							x = _slackx[y];
							break;
						}
						_t[y] = true;
						if (_s[_matchY[y]]) continue;
						q.Enqueue(_matchY[y]);
						AddToTree(_matchY[y], _slackx[y]);
					}
					if (y < _n) break;
				}

				_maxMatch++;

				//инверсируем ребра вдоль дополняющего пути
				int ty;
				for (int cx = x, cy = y; cx != -2; cx = _prev[cx], cy = ty)
				{
					ty = _matchX[cx];
					_matchY[cy] = cx;
					_matchX[cx] = cy;
				}
			}

			#endregion

			return _matchX;
		}

		private void InitMatches()
		{
			for (var i = 0; i < _n; i++)
			{
				_matchX[i] = -1;
				_matchY[i] = -1;
			}
		}

		private void InitSt()
		{
			for (var i = 0; i < _n; i++)
			{
				_s[i] = false;
				_t[i] = false;
			}
		}

		private void InitLbls()
		{
			for (var i = 0; i < _n; i++)
			{
				var minRow = _costMatrix[i, 0];
				for (var j = 0; j < _n; j++)
				{
					if (_costMatrix[i, j] < minRow) minRow = _costMatrix[i, j];
					if (minRow == 0) break;
				}
				_lx[i] = minRow;
			}
			for (var j = 0; j < _n; j++)
			{
				var minColumn = _costMatrix[0, j] - _lx[0];
				for (var i = 0; i < _n; i++)
				{
					if (_costMatrix[i, j] - _lx[i] < minColumn) minColumn = _costMatrix[i, j] - _lx[i];
					if (minColumn == 0) break;
				}
				_ly[j] = minColumn;
			}
		}

		private void UpdateLabels()
		{
			var delta = _inf;
			for (var i = 0; i < _n; i++)
				if (!_t[i])
					if (delta > _slack[i])
						delta = _slack[i];
			for (var i = 0; i < _n; i++)
			{
				if (_s[i])
					_lx[i] = _lx[i] + delta;
				if (_t[i])
					_ly[i] = _ly[i] - delta;
				else _slack[i] = _slack[i] - delta;
			}
		}

		private void AddToTree(int x, int prevx)
		{
			//x-текущая вершина, prevx-вершина от х в альтернативном пути,
			//добавляем ребра (prevx, matchX[x]), (matchX[x],x)

			_s[x] = true; //добавляем x к S
			_prev[x] = prevx;

			var lxx = _lx[x];
			//обновляем slack
			for (var y = 0; y < _n; y++)
			{
				if (_costMatrix[x, y] - lxx - _ly[y] >= _slack[y]) continue;
				_slack[y] = _costMatrix[x, y] - lxx - _ly[y];
				_slackx[y] = x;
			}
		}

		private void InitialMatching()
		{
			for (var x = 0; x < _n; x++)
			{
				for (var y = 0; y < _n; y++)
				{
					if (_costMatrix[x, y] != _lx[x] + _ly[y] || _matchY[y] != -1) continue;
					_matchX[x] = y;
					_matchY[y] = x;
					_maxMatch++;
					break;
				}
			}
		}
	}

        // =======================================================================
        // конец реализации Венгерского алгоритма
        // =======================================================================


        private void button1_Click(object sender, EventArgs e) //открытие файла
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog() {Filter="Формат Excel|*.xls;*.xlsx", ValidateNames=true };


            if (OpenFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                this.textBox_Path.Text = OpenFileDialog1.FileName;
                cboSheet.Enabled = false; //блокируем выбор листа на случай, если решение ещё не содержится в новом файле 
                cboSheet.Text = "Данные";
                button2.Enabled = true;
                //ExeltoGrid();  не требуется т.к. содержится в обработчике cboSheet_SelectedIndexChanged         
            }
        }

        private void button2_Click(object sender, EventArgs e) //решение задачи
        {

            Excel.Application ObjWorkExcel = new Excel.Application(); //открыть эксель
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(textBox_Path.Text
                , Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1 лист
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);//1 ячейку
            //-------------------------------------
            int lastColumn = (int)lastCell.Column;//количество столбцов входного файла
            int lastRow = (int)lastCell.Row;//количество строк
            //-------------------------------------
            String[,] list = new string[lastColumn, lastRow]; // массив значений с листа равен по размеру листу
            for (int i = 0; i < lastColumn; i++) //по всем колонкам
                for (int j = 0; j < lastRow; j++) // по всем строкам
                    list[i, j] = Convert.ToString(((Excel.Range)ObjWorkSheet.Cells[j + 1, i + 1]).Value2);        //считываем текст в строку

            decimal summhoursd = 0;
            //   суммарная нагрузка по всем дисциплинам    
            for (int i = 0; i < lastColumn - 4; i++) //по всем колонкам где записаны часы во 2й строке
                summhoursd = summhoursd + Convert.ToDecimal(list[i + 2, 1]);

            decimal summhoursp = 0;
            //   суммарная нагрузка по всем преподавателям    
            for (int j = 0; j < lastRow - 3; j++) //по всем строкам где записаны часы во 2м столбце
                summhoursp = summhoursp + Convert.ToDecimal(list[1, j + 2]);

            decimal averagehoursd = summhoursd / (lastColumn - 4);
            decimal averagehoursp = summhoursp / averagehoursd;


            var dubl = new int[lastRow - 3]; //считаем количество дублеров для каждого преподавателя

            for (int j = 0; j < lastRow - 3; j++) // по всем строкам  
                dubl[j] = Convert.ToInt32(Math.Round(Convert.ToDecimal(list[1, j + 2]) / averagehoursp));

            int size = 0; // размер матрицы для решения равен сумме дублёров

            for (int j = 0; j < lastRow - 3; j++)   
                size = size + dubl[j];


            var matrix = new int[size, size];
            var shft = 0;

            for (int j = 0; j < lastRow - 3; j++) // по всем преподавателям  
            {
                for (int k = 0; k < dubl[j]; k++) //по всем дублёрам j-го преподавателя
                    for (int i = 0; i < size; i++) //по всем колонкам
                        if (i < lastColumn - 4)
                            matrix[shft + k, i] = 10 - Convert.ToInt32(list[i + 2, j + 2]);
                        else matrix[shft + k, i] = 10;
                //сдвигаем индекс строки на количество дублёров которое мы клонировали j
                shft = shft + dubl[j];
            }

            var algorithm = new HungarianAlgorithm(matrix);

            var result = algorithm.Run();

            String[] names = new string[lastRow - 3];
            //записываем отдельно имена преподавателей
            for (int j = 0; j < lastRow - 3; j++)
                names[j] = list[0, j + 2];

            String[] disciplines = new string[lastColumn - 4];
            //записываем названия дисциплин
            for (int i = 0; i < lastColumn - 4; i++)
                disciplines[i] = list[i + 2, 0];

            shft = 0;
            ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[2];
            //записываем полученное решение на 2й лист входного ехсеl файла

            ((Excel.Range)ObjWorkSheet.Cells[1, 1]).Value2 = "Преподаватели";
            ((Excel.Range)ObjWorkSheet.Cells[1, 2]).Value2 = "Осталось нагрузки";
            ((Excel.Range)ObjWorkSheet.Cells[1, 3]).Value2 = "Предметы";

            for (int j = 0; j < lastRow - 3; j++) // по всем преподавателям  
            {
                ((Excel.Range)ObjWorkSheet.Cells[j + 2, 1]).Value2 = names[j];
                int clmn = 3; //пишем дисциплины в той же строке с третьей колонки
                decimal DiscTime = 0; 
                //загруженность преподавателя j считаем с 0 по всем дисциплинам порученным ему 
                for (int k = 0; k < dubl[j]; k++) //по всем дублёрам j-го преподавателя
                    if (result[shft + k] < disciplines.Length)
                    {
                        ((Excel.Range)ObjWorkSheet.Cells[j + 2, clmn]).Value2 = disciplines[result[shft + k]];
                        DiscTime = DiscTime + Convert.ToDecimal(list[result[shft + k] + 2, 1]);
                        clmn++;
                    }
                ((Excel.Range)ObjWorkSheet.Cells[j + 2, 2]).Value2 = Convert.ToDecimal(list[1, j + 2])-DiscTime;
                shft = shft + dubl[j];
            }

            ObjWorkBook.Close(true, Type.Missing, Type.Missing); //закрыть сохраняя
            ObjWorkExcel.Quit(); // выйти из экселя                                                                                              
            GC.Collect(); // убрать за собой -- в том числе не используемые явно объекты !

            cboSheet.Enabled = true;
            cboSheet.Text = "Решение";
            label1.Visible = true;
            MessageBox.Show("Решение записано на 2й лист Excel файла.");
        }

        private void ExeltoGrid() //вывод из Excel файла на dataGridView требуемого в cboSheet листа
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Refresh();

            Excel.Application ObjWorkExcel = new Excel.Application(); //открыть эксель
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(textBox_Path.Text
                , Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
            int SheetNameIndex = 1;
            if (cboSheet.Text == "Решение") SheetNameIndex = 2;
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[SheetNameIndex]; //получить нужный лист
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);//1 ячейку
            //-------------------------------------
            int lastColumn = (int)lastCell.Column;//количество столбцов входного файла
            int lastRow = (int)lastCell.Row;//количество строк
            //-------------------------------------
            //выводим считанное из файла на экран
            this.dataGridView1.ColumnCount = lastColumn;

            for (int r = 0; r < lastRow; r++)
            {
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(this.dataGridView1);

                for (int c = 0; c < lastColumn; c++)
                    row.Cells[c].Value = Convert.ToString(((Excel.Range)ObjWorkSheet.Cells[r + 1, c + 1]).Value2); 

                this.dataGridView1.Rows.Add(row);
                // выделяем красным цветом отрицательные значения (превышение) изначально запланированной нагрузки в решении
                if ((SheetNameIndex == 2) & (r > 0)) if (Convert.ToDecimal(dataGridView1.Rows[r].Cells[1].Value) < 0 ) 
                        dataGridView1.Rows[r].Cells[1].Style = new DataGridViewCellStyle { ForeColor = Color.Red };

            }
            ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть НЕ сохраняя
            ObjWorkExcel.Quit(); // выйти из экселя
            GC.Collect(); // убрать за собой -- в том числе не используемые явно объекты !

        }

        private void cboSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            ExeltoGrid();
        }
    }
}
