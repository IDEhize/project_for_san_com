using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Project_for_san_com
{
    /// <summary>
    /// Логика взаимодействия для SanComRaportWindow.xaml
    /// </summary>
    public partial class SanComRaportWindow : Window
    {
        public SanComRaportWindow()
        {
            InitializeComponent();
        }

        // Обработчик нажатия на кнопку
        private void AddTextBoxButton_Click(object sender, RoutedEventArgs e)
        {
            // Контейнер для одной строки (пара TextBox)
            StackPanel rowPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(10, 5, 0, 0) // Отступы между строками
            };

            // Первый TextBox
            TextBox textBox1 = new TextBox
            {
                Width = 50, // Ширина TextBox
                Margin = new Thickness(0, 0, 10, 0) // Отступ справа
            };

            // Второй TextBox
            TextBox textBox2 = new TextBox
            {
                Width = 590 // Ширина TextBox
            };

            // Добавляем TextBox в строку
            rowPanel.Children.Add(textBox1);
            rowPanel.Children.Add(textBox2);

            // Добавляем строку в основной контейнер перед кнопкой
            MainStackPanel.Children.Insert(MainStackPanel.Children.Count - 1, rowPanel);
        }
    }
}
