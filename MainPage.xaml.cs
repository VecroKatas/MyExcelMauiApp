using Microsoft.Maui.Controls;
using Microsoft.Maui.Controls.Compatibility;
using System;
using System.Collections;
using System.Collections.Generic;
using Microsoft.Maui.Controls.Shapes;
using Grid = Microsoft.Maui.Controls.Grid;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Globalization;
using Newtonsoft.Json;
using Path = System.IO.Path;

namespace MyExcelMAUIApp
{
    public partial class MainPage : ContentPage
    {
        int CountColumn = 3;
        int CountRow = 3;

        private List<Label> columnLabels;
        private static List<Label> rowLabels;
        private static List<List<CustomCell>> cells;
        private List<List<Entry>> entries;
        private string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "saved.json");
        private Entry previousFocusEntry;
        
        private static bool IsUnary(string str)
        {
            char[] seps = {'+', '-', '*', '/', '^', '%', '(', ')'};
            string[] s = str.Split(seps);
            int k = 0;
            for (int i = 0; i < s.Length; i++)
                if (s[i] != "")
                    k++;
            return (k == 1);
        }

        private static int FindIndexOfPair(string f, char[] brackets, int index)
        {
            int pair = 0;
            Stack<int> stack = new Stack<int>();
            stack.Push(brackets[0]);
            for (int i = index; stack.Count > 0; i++) {
                if (f[i] == brackets[0])
                {
                    stack.Push(i);
                }
                else if (f[i] == brackets[1])
                {
                    if (stack.Any())
                        stack.Pop();
                    pair = i;
                }
            }
            return pair;
        }

        private static string DeleteModDiv(string f, string exp)
        {
            while (f.Contains(exp))
            {
                int start = f.IndexOf(exp) + 4, finish = 0;
                finish = FindIndexOfPair(f, new []{ '(', ')'}, start);

                if (!f.Substring(start, finish - start).Contains(','))
                {
                    throw new Exception("Mod або div не має коми.");
                }
                else
                {
                    string str = f.Substring(start, finish - start);
                    string nr0 = str.Substring(0, str.IndexOf(','));
                    string nr1 = str.Substring(str.IndexOf(',') + 1);
                    string st = exp == "mod(" ? $"({nr0})%({nr1})" : $"({nr0})/({nr1})";
                    f = f.Remove(start - 4, finish - start + 5).Insert(start - 4, st);
                }
            }
            return f;
        }
        private static string Prepare(string f)
        {
            string o = "";
            bool lastIsOp = true;
            
            f = DeleteModDiv(f, "mod(");
            f = DeleteModDiv(f, "div(");
            
            for (int i = 0; i < f.Length; i++)
            {
                if (lastIsOp && f[i] == '-')
                {
                    o += '_';
                }
                else if (!lastIsOp && f[i] == '(')
                {
                    o += "*(";
                }
                else if ("1234567890+-*/^%_()".Contains(f[i]))
                {
                    o += f[i];
                }

                if ("+-*/^%".Contains(f[i]))
                {
                    lastIsOp = true;
                }
                else if ("0123456789_".Contains(f[i]))
                {
                    lastIsOp = false;
                }
            }
            
            return o;
        }
        private static string Calc(string f)
        {
            while(f.Contains('(') && f.Contains(')'))
            {
                int start = f.IndexOf('(');
                int finish = FindIndexOfPair(f, new[] {'(', ')'}, start + 1);
                f = f.Substring(0, f.IndexOf('(')) + Calc(f.Substring(start + 1, 
                    finish - start - 1)) + f.Substring(finish + 1);
            }
            
            f = Prepare(f);

            while (f.Contains('^'))
            {
                f = Pow(f);
            }
            while (f.Contains('*') || f.Contains('/'))
            {
                if(f.IndexOf('*') >= 0 && (f.IndexOf('*') < f.IndexOf('/') || f.IndexOf('/') < 0))
                {
                    f = Mult(f);
                }
                else
                {
                    f = Div(f);
                }
            }
            while (f.Contains('%'))
            {
                f = Mod(f);
            }
            while (f.Contains('+') || f.Contains('-'))
            {
                if (f.IndexOf('+') >= 0 && (f.IndexOf('+') < f.IndexOf('-') || f.IndexOf('-') < 0))
                {
                    f = Add(f);
                }
                else
                {
                    f = Sub(f);
                }
            }

            return f.Replace('_', '-');
        }
        private static string Mult(string s)
        {
            if (!s.Contains('*')) return s;

            s = s.Replace(" ", "");
            s = s.Replace(",", ".");

            int opi = s.IndexOf('*');
            string nr0 = "";
            string nr1 = "";
            for(int i=opi-1; i>=0 && "1234567890._".Contains(s[i]); i--)
            {
                nr0 = s[i] + nr0;
            }
            for (int i = opi + 1; i < s.Length && "1234567890._".Contains(s[i]); i++)
            {
                nr1 += s[i];
            }

            nr0 = nr0.Replace('_', '-');
            nr1 = nr1.Replace('_', '-');
            string res = (int.Parse(nr0) * int.Parse(nr1)).ToString().Replace('-', '_');

            return s.Substring(0, opi - nr0.Length) + res + s.Substring(opi + 1 + nr1.Length);
        }
        private static string Pow(string s)
        {
            if (!s.Contains('^')) return s;

            s = s.Replace(" ", "");
            s = s.Replace(",", ".");

            int opi = s.IndexOf('^');
            string nr0 = "";
            string nr1 = "";
            for (int i = opi - 1; i >= 0 && "1234567890._".Contains(s[i]); i--)
            {
                nr0 = s[i] + nr0;
            }
            for (int i = opi + 1; i < s.Length && "1234567890._".Contains(s[i]); i++)
            {
                nr1 += s[i];
            }

            nr0 = nr0.Replace('_', '-');
            nr1 = nr1.Replace('_', '-');
            string res = Convert.ToInt32(Math.Pow(int.Parse(nr0), int.Parse(nr1))).ToString().Replace('-', '_');

            return s.Substring(0, opi - nr0.Length) + res + s.Substring(opi + 1 + nr1.Length);
        }
        private static string Mod(string s)
        {
            if (!s.Contains('%')) return s;

            s = s.Replace(" ", "");
            s = s.Replace(",", ".");

            int opi = s.IndexOf('%');
            string nr0 = "";
            string nr1 = "";
            for (int i = opi - 1; i >= 0 && "1234567890._".Contains(s[i]); i--)
            {
                nr0 = s[i] + nr0;
            }
            for (int i = opi + 1; i < s.Length && "1234567890._".Contains(s[i]); i++)
            {
                nr1 += s[i];
            }

            nr0 = nr0.Replace('_', '-');
            nr1 = nr1.Replace('_', '-');
            string res = (int.Parse(nr0) % int.Parse(nr1)).ToString().Replace('-', '_');
            
            return s.Substring(0, opi - nr0.Length) + res + s.Substring(opi + 1 + nr1.Length);
        }
        private static string Div(string s)
        {
            if (!s.Contains('/')) return s;

            s = s.Replace(" ", "");
            s = s.Replace(",", ".");

            int opi = s.IndexOf('/');
            string nr0 = "";
            string nr1 = "";
            for (int i = opi - 1; i >= 0 && "1234567890._".Contains(s[i]); i--)
            {
                nr0 = s[i] + nr0;
            }
            for (int i = opi + 1; i < s.Length && "1234567890._".Contains(s[i]); i++)
            {
                nr1 += s[i];
            }

            nr0 = nr0.Replace('_', '-');
            nr1 = nr1.Replace('_', '-');
            string res = (int.Parse(nr0) / int.Parse(nr1)).ToString().Replace('-', '_');
            
            return s.Substring(0, opi - nr0.Length) + res + s.Substring(opi + 1 + nr1.Length);
        }
        private static string Add(string s)
        {
            if (!s.Contains('+')) return s;

            s = s.Replace(" ", "");
            s = s.Replace(",", ".");

            int opi = s.IndexOf('+');
            string nr0 = "";
            string nr1 = "";
            for (int i = opi - 1; i >= 0 && "1234567890._".Contains(s[i]); i--)
            {
                nr0 = s[i] + nr0;
            }
            for (int i = opi + 1; i < s.Length && "1234567890._".Contains(s[i]); i++)
            {
                nr1 += s[i];
            }

            nr0 = nr0.Replace('_', '-');
            nr1 = nr1.Replace('_', '-');
            string res = (int.Parse(nr0) + int.Parse(nr1)).ToString().Replace('-', '_');
            
            return s.Substring(0, opi - nr0.Length) + res + s.Substring(opi + 1 + nr1.Length);
        }
        private static string Sub(string s)
        {
            if (!s.Contains('-')) return s;

            s = s.Replace(" ", "");
            s = s.Replace(",", ".");

            int opi = s.IndexOf('-');
            string nr0 = "";
            string nr1 = "";
            for (int i = opi - 1; i >= 0 && "1234567890._".Contains(s[i]); i--)
            {
                nr0 = s[i] + nr0;
            }
            for (int i = opi + 1; i < s.Length && "1234567890._".Contains(s[i]); i++)
            {
                nr1 += s[i];
            }

            nr0 = nr0.Replace('_', '-');
            nr1 = nr1.Replace('_', '-');
            string res = (int.Parse(nr0) - int.Parse(nr1)).ToString().Replace('-', '_');
            
            return s.Substring(0, opi - nr0.Length) + res + s.Substring(opi + 1 + nr1.Length);
        }

        public class NewGrid
        {
            public List<List<CustomCell>> NewCells { get; set; }
        }
        
        public class CustomCell
        {
            [JsonProperty]
            public int Number { get; set; }
            //public Expression Expression { get; set; }
            [JsonProperty]
            public string ExpressionString { get; set; }
            [JsonProperty]
            public string Value { get; set; }

            public CustomCell(int number, string expression)
            {
                Number = number;
                CreateExpression(expression);
                //Expression = new Expression(expression);
            }

            public void CreateExpression(string expression)
            {
                if (expression[0] == '=')
                {
                    ExpressionString = expression;
                    expression = ChangeCells(expression);
                    CheckExpression(expression);
                }
                else
                {
                    ExpressionString = expression;
                    Value = expression;
                }
            }
            private string ChangeCells(string expression)
            {
                string s = "";
                int start = 0, finish = 0;
                while (expression.Contains('{'))
                {
                    start = expression.IndexOf('{');
                    finish = FindIndexOfPair(expression, new[] {'{', '}'}, start + 1);
                    string[] name = expression.Substring(start + 1, finish - start - 1).Split(':');
                    s = cells[GetColumnIndex(name[0])][Convert.ToInt32(name[1]) - 1].Value;
                    expression = expression.Substring(0, start) + s + expression.Substring(finish + 1);
                }
                return expression;
            }

            private int GetColumnIndex(string name)
            {
                int number = 0;
                for (int i = 0; i < name.Length; i++)
                {
                    number += (name[name.Length - i - 1] - 64) * (int)Math.Pow(26, i);
                }
                return number - 1;
            }

            private void CheckExpression(string f)
            {
                char[] seps = {'+', '-', '*', '/', '^', '%', '(', ')', '='};
                if (f.Contains("--") || f.Contains("++") && IsUnary(f))
                {
                    if (f.Contains("--"))
                        f = (Convert.ToInt32(f.Split(seps)[1]) - 1).ToString();
                    else if (f.Contains("++"))
                        f = (Convert.ToInt32(f.Split(seps)[1]) + 1).ToString();
                    Value = f;
                }
                else
                {
                    CalculateValue(f);
                }
            }
            
            private void CalculateValue(string expression)
            {
                string expr = "";
                try
                {
                    expr = Prepare(expression);
                    Value = Calc(expr);
                }
                catch (Exception e)
                {
                    Value = expr;
                    throw;
                }
            }
        }
        
        public class Expression
        {
            public string ExpressionString { get; set; }
            public string Value { get; set; }

            public Expression(string expression)
            {
                if (expression[0] == '=')
                {
                    ExpressionString = expression;
                    expression = ChangeCells(expression);
                    CheckExpression(expression);
                }
                else
                {
                    ExpressionString = expression;
                    Value = expression;
                }
            }

            private string ChangeCells(string expression)
            {
                string s = "";
                int start = 0, finish = 0;
                while (expression.Contains('{'))
                {
                    start = expression.IndexOf('{');
                    finish = FindIndexOfPair(expression, new[] {'{', '}'}, start + 1);
                    string[] name = expression.Substring(start + 1, finish - start - 1).Split(':');
                    //s = cells[GetColumnIndex(name[0])][Convert.ToInt32(name[1]) - 1].Expression.Value;
                    expression = expression.Substring(0, start) + s + expression.Substring(finish + 1);
                }
                return expression;
            }

            private int GetColumnIndex(string name)
            {
                int number = 0;
                for (int i = 0; i < name.Length; i++)
                {
                    number += (name[name.Length - i - 1] - 64) * (int)Math.Pow(26, i);
                }
                return number - 1;
            }

            private void CheckExpression(string f)
            {
                char[] seps = {'+', '-', '*', '/', '^', '%', '(', ')', '='};
                if (f.Contains("--") || f.Contains("++") && IsUnary(f))
                {
                    if (f.Contains("--"))
                        f = (Convert.ToInt32(f.Split(seps)[1]) - 1).ToString();
                    else if (f.Contains("++"))
                        f = (Convert.ToInt32(f.Split(seps)[1]) + 1).ToString();
                    Value = f;
                }
                else
                {
                    CalculateValue(f);
                }
            }
            
            private void CalculateValue(string expression)
            {
                string expr = "";
                try
                {
                    expr = Prepare(expression);
                    Value = Calc(expr);
                }
                catch (Exception e)
                {
                    Value = expr;
                    throw;
                }
            }
        }
        
        public MainPage()
        {
            cells = new List<List<CustomCell>>();
            entries = new List<List<Entry>>();
            columnLabels = new List<Label>();
            rowLabels = new List<Label>();
            InitializeComponent();
            CreateGrid();
            
            textInput.Unfocused += TextInput_Unfocused;
        }
        
        private async void CreateGrid()
        {
            try
            {
                AddColumnsAndColumnLabels();
                AddRowsAndCellEntries();

            }
            catch (Exception e)
            {
                await DisplayAlert("CreateGrid", e.Message, "damn");
                throw;
            }
        }

        private void AddColumnsAndColumnLabels()
        {
            for (int col = 0; col < CountColumn + 1; col++)
            {
                grid.ColumnDefinitions.Add(new ColumnDefinition());
                if (col > 0)
                {
                    if (cells.Count < col)
                        cells.Add(new List<CustomCell>());
                    entries.Add(new List<Entry>());
                    var label = new Label
                    {
                        Text = GetColumnName(col) + " " + grid.Children.Count,
                        VerticalOptions = LayoutOptions.Center,
                        HorizontalOptions = LayoutOptions.Center
                    };
                    Grid.SetRow(label, 0);
                    Grid.SetColumn(label, col);
                    grid.Children.Add(label);
                    columnLabels.Add(label);
                }
            }
        }

        private void AddRowsAndCellEntries()
        {
            grid.RowDefinitions.Add(new RowDefinition());
            for (int row = 0; row < CountRow; row++)
            {
                grid.RowDefinitions.Add(new RowDefinition());
                var label = new Label
                {
                    Text = (row + 1).ToString()  + " " + grid.Children.Count,
                    VerticalOptions = LayoutOptions.Center,
                    HorizontalOptions = LayoutOptions.Center
                };
                Grid.SetRow(label, row + 1);
                Grid.SetColumn(label, 0);
                grid.Children.Add(label);
                rowLabels.Add(label);
                for (int col = 0; col < CountColumn; col++)
                {
                    var entry = new Entry
                    {
                        Text = "" + grid.Children.Count,
                        VerticalOptions = LayoutOptions.Center,
                        HorizontalOptions = LayoutOptions.Center
                    };
                    
                    if (cells[col].Count == columnLabels.Count)
                    {
                        entry.Text = cells[col][row].Value;
                    }
                    else
                    {
                        cells[col].Add(new CustomCell(grid.Children.Count, entry.Text));
                    }
                    
                    entry.Focused += Entry_Focused;
                    entry.TextChanged += Entry_TextChanged;
                    entry.Unfocused += Entry_Unfocused;
                    Grid.SetRow(entry, row + 1);
                    Grid.SetColumn(entry, col + 1);
                    entries[col].Add(entry);
                    grid.Children.Add(entry);
                }
            }
        }

        private string GetColumnName(int colIndex)
        {
            int dividend = colIndex;
            string columnName = string.Empty;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }

        private void Entry_Focused(object sender, FocusEventArgs e)
        {
            var entry = (Entry) sender;
            var row = Grid.GetRow(entry) - 1;
            var col = Grid.GetColumn(entry) - 1;
            var content = entry.Text;
            CustomCell cell = cells[col][row];
            //if (cell.Expression != null)
            if (cell.ExpressionString != null)
            {
                //textInput.Text = cell.Expression.ExpressionString;
                //entries[col][row].Text = cell.Expression.Value.ToString();
                textInput.Text = cell.ExpressionString;
                entries[col][row].Text = cell.Value.ToString();
            }
            else
            {
                textInput.Text = content;
                entries[col][row].Text = content;
            }
        }

        private void Entry_TextChanged(object sender, EventArgs e)
        {
            var entry = (Entry) sender;
            var row = Grid.GetRow(entry) - 1;
            var col = Grid.GetColumn(entry) - 1;
            var content = entry.Text;
            if (entry.IsFocused)
            {
                textInput.Text = content;
                entries[col][row].Text = content;
            }
        }

        private void Entry_Unfocused(object sender, FocusEventArgs e)
        {
            var entry = (Entry) sender;
            var row = Grid.GetRow(entry) - 1;
            var col = Grid.GetColumn(entry) - 1;
            var content = entry.Text;
            ChangeCell(col, row, content);
            CustomCell cell = cells[col][row];
            /*if (cell.Expression == null || (cell.Expression.ExpressionString != content.Substring(1) && content[0] == '='))
                cell.Expression = new Expression(content);
            textInput.Text = cell.Expression.ExpressionString;
            entries[col][row].Text = cell.Expression.Value.ToString();*/
            previousFocusEntry = entry;
        }
        
        private void TextInput_Unfocused(object sender, FocusEventArgs e)
        {
            var entry = new Entry();
            if (previousFocusEntry != null)
                entry = previousFocusEntry;
            else
                entry = entries[0][0];
            var row = Grid.GetRow(entry) - 1;
            var col = Grid.GetColumn(entry) - 1;
            var content  = textInput.Text;
            ChangeCell(col, row, content);
        }

        private async void ChangeCell(int col, int row, string content)
        {
            try
            {
                CustomCell cell = cells[col][row];
                /*cell.Expression = new Expression(content);
                textInput.Text = cell.Expression.ExpressionString;
                entries[col][row].Text = cell.Expression.Value;*/
                cell.CreateExpression(content);
                textInput.Text = cell.ExpressionString;
                entries[col][row].Text = cell.Value;
            }
            catch (Exception e)
            {
                await DisplayAlert("Помилка", e.ToString(), "Ok");
                throw;
            }
        }
        
        private void CalculateButton_Clicked(object sender, EventArgs e)
        {
            
        }

        private void SaveButton_Clicked(object sender, EventArgs e)
        {
            Save(cells);
        }

        private void Save<T>(T serializableObject)
        {
            try
            {
                string json = System.Text.Json.JsonSerializer.Serialize<T>(serializableObject);;
                textInput.Text = path;
                File.WriteAllText(path, json);
            }
            catch (Exception exception)
            {
                DisplayAlert("Помилка", "Шось відбулося: " + exception.Message, "Ok");
                throw;
            }
        }

        private async void ReadButton_Clicked(object sender, EventArgs e)
        {
            string json = File.ReadAllText(path);
            try
            {
                var newCells = JsonConvert.DeserializeObject<List<List<CustomCell>>>(json);
                /*CountColumn = newCells.Length;
                CountRow = newCells[0].Length;
                foreach (var list in entries)
                    list.Clear();
                foreach (var list in cells)
                    list.Clear();
                entries.Clear();
                cells.Clear();
                grid.Children.Clear();
                columnLabels.Clear();
                rowLabels.Clear();
                CountColumn = newCells.Length;
                CountRow = newCells[0].Length;
                
                for (int i = 0; i < newCells.Length; i++)
                    cells.Add(new List<CustomCell>());
                for (int i = 0; i < newCells.Length; i++)
                {
                    for (int j = 0; j < newCells[0].Length; j++)
                    {
                        cells[i].Add(new CustomCell(0, newCells[i][j].ExpressionString));
                    }
                }
            
                CreateGrid();*/
            }
            catch (Exception exception)
            {
                await DisplayAlert("Load", json, "Ok");
                await DisplayAlert("Load", exception.Message, "Ok");
            }
        }
        
        static T Load<T>(string filepath)
        {
            string json = File.ReadAllText(filepath);
            T serializableObject = JsonConvert.DeserializeObject<T>(json);
            return serializableObject;
        }

        private async void ExitButton_Clicked(object sender, EventArgs e)
        {
            bool answer = await DisplayAlert("Підтвердження", "Ви дійсно хочете вийти?", "Так", "Ні");
            if (answer)
            {
                System.Environment.Exit(0);
            }
        }

        private async void HelpButton_Clicked(object sender, EventArgs e)
        {
            await DisplayAlert("Довідка", "Лабораторна робота 1. Студента Катасонова Тимура К25", "OK");
        }

        private void DeleteRowButton_Clicked(object sender, EventArgs e)
        {
            if (grid.RowDefinitions.Count > 1)
            {
                int lastRowIndex = grid.RowDefinitions.Count - 1;
                grid.RowDefinitions.RemoveAt(lastRowIndex);
                grid.Children.RemoveAt(grid.Children.IndexOf(rowLabels[rowLabels.Count - 1]));
                for (int col = 0; col < CountColumn; col++)
                {
                    //RefillCells(grid.Children.IndexOf(entries[col][lastRowIndex - 1]));
                    
                    grid.Children.RemoveAt(grid.Children.IndexOf(entries[col][lastRowIndex - 1])); 
                }

                for (int col = 0; col < CountColumn; col++)
                    cells[col].RemoveAt(cells[col].Count - 1);
                for (int col = 0; col < CountColumn; col++)
                    entries[col].RemoveAt(entries[col].Count - 1);
                rowLabels.RemoveAt(rowLabels.Count - 1);
                CountRow--;
                //RefreshCells();
            }
        }
        
        private void DeleteColumnButton_Clicked(object sender, EventArgs e)
        {
            if (grid.ColumnDefinitions.Count > 1)
            {
                int lastColumnIndex = grid.ColumnDefinitions.Count - 1;
                grid.ColumnDefinitions.RemoveAt(lastColumnIndex);
                //RefillCells(grid.Children.IndexOf(columnLabels[columnLabels.Count - 1]));
                grid.Children.RemoveAt(grid.Children.IndexOf(columnLabels[columnLabels.Count - 1]));
                for (int row = 0; row < CountRow; row++)
                {
                    //RefillCells(grid.Children.IndexOf(entries[lastColumnIndex - 1][row]));
                    
                    grid.Children.RemoveAt(grid.Children.IndexOf(entries[lastColumnIndex - 1][row]));
                }
                
                while(cells[lastColumnIndex - 1].Count > 0)
                    cells[lastColumnIndex - 1].RemoveAt(0);
                while(entries[lastColumnIndex - 1].Count > 0)
                    entries[lastColumnIndex - 1].RemoveAt(0);
                columnLabels.RemoveAt(columnLabels.Count - 1);
                CountColumn--;
                //RefreshCells();
            }
        }

        private void RefreshCells()
        {
            foreach (var col in entries)
            {
                foreach (var entry in col)
                {
                    entry.Text = cells[Grid.GetColumn(entry) - 1][Grid.GetRow(entry) - 1].Number.ToString();
                }
            }
        }

        private void RefillCells(int startIndex)
        {
            for (int col = 0; col < CountColumn; col++)
            {
                for (int row = 0; row < CountRow; row++)
                {
                    cells[col][row].Number = grid.Children.IndexOf(entries[col][row]);
                }
            }
        }

        private void AddRowButton_Clicked(object sender, EventArgs e)
        {
            AddRow();
        }

        private void AddColumnButton_Clicked(object sender, EventArgs e)
        {
            int newColumn = grid.ColumnDefinitions.Count;
            grid.ColumnDefinitions.Add(new ColumnDefinition());
            var label = new Label
            {
                Text = GetColumnName(newColumn) + " " + grid.Children.Count,
                VerticalOptions = LayoutOptions.Center,
                HorizontalOptions = LayoutOptions.Center
            };
            Grid.SetRow(label, 0);
            Grid.SetColumn(label, newColumn);
            grid.Children.Add(label);
            columnLabels.Add(label);
            cells.Add(new List<CustomCell>());
            entries.Add(new List<Entry>());
            for (int row = 0; row < CountRow; row++)
            {
                var entry = new Entry
                {
                    Text = "" + grid.Children.Count,
                    VerticalOptions = LayoutOptions.Center,
                    HorizontalOptions = LayoutOptions.Center
                };
                entry.Focused += Entry_Focused;
                entry.TextChanged += Entry_TextChanged;
                entry.Unfocused += Entry_Unfocused;
                Grid.SetRow(entry, row + 1);
                Grid.SetColumn(entry, newColumn);
                cells[newColumn - 1].Add(new CustomCell(grid.Children.Count, entry.Text));
                grid.Children.Add(entry);
                entries[newColumn - 1].Add(entry);
            }

            CountColumn++;
        }
        
        private void AddRow()
        {
            int newRow = grid.RowDefinitions.Count;
            grid.RowDefinitions.Add(new RowDefinition());
            var label = new Label
            {
                Text = newRow.ToString() + " " + grid.Children.Count,
                VerticalOptions = LayoutOptions.Center,
                HorizontalOptions = LayoutOptions.Center
            };
            Grid.SetRow(label, newRow);
            Grid.SetColumn(label, 0);
            grid.Children.Add(label);
            rowLabels.Add(label);
            
            for (int col = 0; col < CountColumn; col++)
            {
                var entry = new Entry
                {
                    Text = "" + grid.Children.Count,
                    VerticalOptions = LayoutOptions.Center,
                    HorizontalOptions = LayoutOptions.Center
                };
                entry.Focused += Entry_Focused;
                entry.TextChanged += Entry_TextChanged;
                entry.Unfocused += Entry_Unfocused;
                Grid.SetRow(entry, newRow);
                Grid.SetColumn(entry, col + 1);
                cells[col].Add(new CustomCell(grid.Children.Count, entry.Text));
                grid.Children.Add(entry);
                entries[col].Add(entry);
            }
            CountRow++;
        }
    }
}