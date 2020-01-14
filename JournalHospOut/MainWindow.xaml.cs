using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data.OleDb;
using System.Data;
using System.Configuration;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Text.RegularExpressions;

namespace JournalHospOut
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window,INotifyPropertyChanged
    {
        private ICollectionView _mkbCollection;//
        private string _filterString;//
        public OleDbConnection cn = new OleDbConnection();
        OleDbCommand cmd = new OleDbCommand();
        OleDbDataReader dr;
        DataTable dtForCombo = new DataTable();
        
        List<mkb10> mkbList = new List<mkb10>();

        string conPath = ConfigurationManager.AppSettings.Get("conPath");
        string rowCountOnList = ConfigurationManager.AppSettings.Get("rowCountOnList");
        public MainWindow()
        {
            InitializeComponent();
            cn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + conPath;
            cmd.Connection = cn;
            mkbCollection = CollectionViewSource.GetDefaultView(mkbList);//
            mkbCollection.Filter = new Predicate<object>(Filter);
            tbCountRowShow.Text = rowCountOnList;
            
        }

        public ICollectionView mkbCollection//
        {
            get { return _mkbCollection; }//
            set { _mkbCollection = value; NotifyPropertyChanged("mkbCollection"); }//
        }

        public string FilterString//
        {
            get { return _filterString; }//
            set//
            {
                _filterString = value;//
                NotifyPropertyChanged("FilterString");//
                FilterCollection();//
            }
        }

        private void FilterCollection()//
        {
            if (_mkbCollection != null)//
            {
                _mkbCollection.Refresh();//
            }
        }

        public bool Filter(object obj)//
        {
            var data = obj as mkb10;//
            if(data != null)//
            {
                if (!string.IsNullOrEmpty(_filterString))//
                {
                    return data.mkb_b.Contains(_filterString) || data.mkb_t.Contains(_filterString);//
                }
                return true;//
            }
            return false;//
        }

        public event PropertyChangedEventHandler PropertyChanged;//
        private void NotifyPropertyChanged(string property)//
        {
            if (PropertyChanged != null)//
            {
                PropertyChanged(this, new PropertyChangedEventArgs(property));//
            }
        }

        int lastList()
        {
            int maxList = 1;
            cmd.CommandText = @"select count(nom) from jho";
            try
            {
                cn.Open();
                maxList = (int)cmd.ExecuteScalar() / Convert.ToUInt16(tbCountRowShow.Text)+1;
                cn.Close();
                if (maxList < 1)
                    maxList = 1;
            }
            catch
            {
                cn.Close();
                maxList = 1;
            }
            return maxList;
        }

        private void loadDataGrid()
        {
            try
            {
                int x = (Convert.ToUInt16(tbListJournal.Text) - 1) * Convert.ToUInt16(tbCountRowShow.Text);
                if (x < 1)
                    cmd.CommandText = @"select top " + tbCountRowShow.Text + " * from jho order by nom ASC";
                else
                    cmd.CommandText = @"select top " + tbCountRowShow.Text + " * from jho where nom > "+ x +" order by nom ASC";
                //int x = (Convert.ToUInt16(tbListJournal.Text) - 1) * Convert.ToUInt16(tbCountRowShow.Text);
                //if (x < 1)
                //    cmd.CommandText = @"select top " + tbCountRowShow.Text + " * from jho order by nom DESC";
                //else
                //    cmd.CommandText = @"select top " + tbCountRowShow.Text + " * from jho where nom not in (select top " + x + " nom from jho order by nom DESC) order by nom DESC";
                cn.Open();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                DataTable dtForMainGrid = new DataTable();
                da.Fill(dtForMainGrid);
                cn.Close();
                dataGrid.ItemsSource = dtForMainGrid.DefaultView;
                
            }
            catch (Exception ex)
            {
                cn.Close();
                MessageBox.Show("Error" + ex);
            }
            tbMkb.Focus();
        }
        private void loadMKBData()
        {
            try
            {
                cn.Open();
                cmd.CommandText = "select distinct mkb_b, mkb_t from mkb";
                dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    mkbList.Add(new mkb10
                    {
                        mkb_b = dr[0].ToString(),
                        mkb_t = dr[1].ToString()
                    });
                }
                dr.Close();
                cn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error" + ex);
            }
        }

        private void proverki()
        {
            try
            {
                Convert.ToUInt16(tbListJournal.Text);
            }
            catch
            {
                tbListJournal.Text = "1";
            }
            try
            {
                Convert.ToUInt16(tbCountRowShow.Text);
            }
            catch
            {
                tbCountRowShow.Text = rowCountOnList;
            }
            try
            {
                Convert.ToUInt16(tbNom.Text);
            }
            catch
            {
                setNom();
            }
            try
            {
                Convert.ToUInt16(tbKd.Text);
            }
            catch
            {
                tbKd.Clear();
            }
            try
            {
                Convert.ToUInt16(tbAge.Text);
            }
            catch
            {
                tbAge.Clear();
            }
        }
        private void UI_Loaded(object sender, RoutedEventArgs e)
        {
            tbListJournal.Text = lastList().ToString();
            loadDataGrid();
            loadMKBData();
            setNom();
        }

        private void clearField()
        {
            tbNom.Clear();
            tbMkb.Clear();
            tbKd.Clear();
            rbFemale.IsChecked = false;
            rbMale.IsChecked = false;
            tbAge.Clear();
            setNom();
        }

        private void insertData()
        {
            string _pol=null;
            proverki();
            if (rbMale.IsChecked.Value)
            {
                _pol = "М";
            }
            if (rbFemale.IsChecked.Value)
            {
                _pol = "Ж";
            }
            if (tbNom.Text != "" & tbMkb.Text != "" & tbKd.Text != "" & tbAge.Text != "" & _pol != null)
            {
                cmd.CommandText = "insert into jho (nom, mkb, kd, pol, age) values('"+tbNom.Text+"', '"+tbMkb.Text+"', '"+tbKd.Text+"', '"+_pol+"', '"+tbAge.Text+"')";
                try
                {
                    cn.Open();
                    cmd.ExecuteNonQuery();
                    cn.Close();
                    int tmp = lastList();
                    if (Convert.ToUInt16(tbListJournal.Text) > tmp || Convert.ToUInt16(tbListJournal.Text) < tmp)
                        tbListJournal.Text = tmp.ToString();
                    loadDataGrid();
                    clearField();
                }
                catch (Exception e)
                {
                    cn.Close();
                    MessageBox.Show("" + e);
                }
                
            }
            else
            {
                MessageBox.Show("Заполнены не все поля");
            }
        }

        private void updateData()
        {
            string _pol = null;
            proverki();
            if (rbMale.IsChecked.Value)
            {
                _pol = "М";
            }
            if (rbFemale.IsChecked.Value)
            {
                _pol = "Ж";
            }
            if (tbNom.Text != "" & tbMkb.Text != "" & tbKd.Text != "" & tbAge.Text != "" & _pol != null)
            {
                if (MessageBox.Show("Изменить запись?", "", MessageBoxButton.OKCancel) == MessageBoxResult.OK)
                {
                    cmd.CommandText = "update jho set mkb='" + tbMkb.Text + "' ,kd='" + tbKd.Text + "' ,pol='" + _pol + "' ,age='" + tbAge.Text + "' where nom = " + tbNom.Text;
                    try
                    {
                        cn.Open();
                        cmd.ExecuteNonQuery();
                        cn.Close();
                        loadDataGrid();
                        clearField();
                    }
                    catch (Exception e)
                    {
                        cn.Close();
                        MessageBox.Show("" + e);
                    }
                }
                else
                {
                    MessageBox.Show("Заполнены не все поля");
                }
            }
        }

        private void btDelete_Click(object sender, RoutedEventArgs e)
        {
            if (tbNom.Text != null)
            {
                if (MessageBox.Show("Удалить запись?", "", MessageBoxButton.OKCancel) == MessageBoxResult.OK)
                {
                    try
                    {
                        cmd.CommandText = "delete from jho where nom="+tbNom.Text;
                        cn.Open();
                        cmd.ExecuteNonQuery();
                        cn.Close();
                        loadDataGrid();
                        clearField();
                    }
                    catch (Exception ex)
                    {
                        cn.Close();
                        MessageBox.Show(""+ex);
                    }
                }
            }
        }

        private void setNom()
        {
            cmd.CommandText = "select MAX(nom) from jho";
            cn.Open();
            try
            {
                tbNom.Text = ((int)cmd.ExecuteScalar() + 1).ToString();
            }
            catch
            {
                tbNom.Text = "1";
            }
            cn.Close();
        }
        private void dataGrid_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            if (e.Column.Header.ToString() == "nom")
                e.Column.Header = "Номер";
            if (e.Column.Header.ToString() == "mkb")
                e.Column.Header = "МКБ";
            if (e.Column.Header.ToString() == "kd")
                e.Column.Header = "К/д";
            if (e.Column.Header.ToString() == "pol")
                e.Column.Header = "Пол";
            if (e.Column.Header.ToString() == "age")
                e.Column.Header = "Возраст";
        }

        private void dataGrid_AutoGeneratedColumns(object sender, EventArgs e)
        {
            dataGrid.Items.SortDescriptions.Add(new SortDescription(dataGrid.Columns[0].SortMemberPath, ListSortDirection.Ascending));
        }

        private void bnPlusList_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Convert.ToUInt16(tbListJournal.Text);
            }
            catch
            {
                tbListJournal.Text = "1";
            }
            tbListJournal.Text = (Convert.ToUInt16(tbListJournal.Text) + 1).ToString();
            loadDataGrid();
        }

        private void bnMinusList_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Convert.ToUInt16(tbListJournal.Text);
            }
            catch
            {
                tbListJournal.Text = "1";
            }
            if (Convert.ToUInt16(tbListJournal.Text) > 1)
            {
                tbListJournal.Text = (Convert.ToUInt16(tbListJournal.Text) - 1).ToString();
                loadDataGrid();
            }
        }

        private void tbCountRowShow_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                bnRefresh.Focus();
            }
        }

        private void bnRefresh_Click(object sender, RoutedEventArgs e)
        {
            proverki();
            int tmp = lastList();
            if (Convert.ToUInt16(tbListJournal.Text) > tmp)
                tbListJournal.Text = tmp.ToString();
            loadDataGrid();
        }

        private void tbListJournal_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                bnRefresh.Focus();
            }
        }

        private void button_Copy1_Click(object sender, RoutedEventArgs e)
        {
            clearField();
            setNom();
        }

        private void btSaveData_Click(object sender, RoutedEventArgs e)
        {
            insertData();
            
        }

        private void tbNom_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Regex regex = new Regex("[^0-9]+");
            e.Handled = regex.IsMatch(e.Text);
        }

        private void tbKd_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Regex regex = new Regex("[^0-9]+");
            e.Handled = regex.IsMatch(e.Text);
        }

        private void tbAge_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Regex regex = new Regex("[^0-9]+");
            e.Handled = regex.IsMatch(e.Text);
        }

        private void tbListJournal_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Regex regex = new Regex("[^0-9]+");
            e.Handled = regex.IsMatch(e.Text);
        }

        private void tbCountRowShow_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Regex regex = new Regex("[^0-9]+");
            e.Handled = regex.IsMatch(e.Text);
        }

        private void dataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                DataRowView row = (DataRowView)(dataGrid.SelectedItems[0]);
                tbNom.Text = row[0].ToString();
                tbMkb.Text = row[1].ToString();
                tbKd.Text = row[2].ToString();
                if (row[3].ToString() == "М")
                    rbMale.IsChecked = true;
                if (row[3].ToString() == "Ж")
                    rbFemale.IsChecked = true;
                tbAge.Text = row[4].ToString();
            }
            catch //(Exception ex)
            {
                //MessageBox.Show("" + ex);
                clearField();
            }
        }

        private void dgMkb_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                mkb10 row1 = (mkb10)dgMkb.SelectedItem;
                tbMkb.Text = row1.mkb_b;
                tbKd.Focus();
            }
            catch //(Exception ex)
            {
                //MessageBox.Show("" + ex);
            }
        }

        private void btUpdateData_Click(object sender, RoutedEventArgs e)
        {
            updateData();
        }

        private void tbNom_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                tbMkb.Focus();
        }

        private void tbMkb_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                tbKd.Focus();
        }

        private void tbKd_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                rbMale.Focus();
        }

        private void tbAge_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                btSaveData.Focus();
        }

        private void rbMale_Checked(object sender, RoutedEventArgs e)
        {
            tbAge.Focus();
        }

        private void rbFemale_Checked(object sender, RoutedEventArgs e)
        {
            tbAge.Focus();
        }

        private void btExcel_Click(object sender, RoutedEventArgs e)
        {
            test();
        }

        private void tbListJournal_TextChanged(object sender, TextChangedEventArgs e)
        {
            int tmp = lastList();
            if (Convert.ToUInt16(tbListJournal.Text) > tmp)
                tbListJournal.Text = tmp.ToString();
        }

        private void bnMinusList_Copy_Click(object sender, RoutedEventArgs e)
        {
            if (tbListJournal.Text != "1")
            {
                tbListJournal.Text = "1";
                loadDataGrid();
            }
        }

        private void bnPlusList_Copy_Click(object sender, RoutedEventArgs e)
        {
            int tmp = lastList();
            if (tbListJournal.Text != tmp.ToString())
            {
                tbListJournal.Text = tmp.ToString();
                loadDataGrid();
            }
        }
    }

    public class mkb10
    {
        public string mkb_b { get; set; }
        public string mkb_t { get; set; }
    }
}
