   M S B A M L   `   `   `  ���� A  =Parser, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullclr-namespace:ParserParser      Parser.MainWindow   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     &localclr-namespace:Parser       � Title$   ��$	��900��$
��1600��$	��450��$	��800�� � Icon$ Resources/iconParse.ico��.���  � RowDefinitions �� $��50 ��� $��8* �.+��� $��0�����#28719e� $��0�� � ColumnDefinitions �� $
��1.5* ��� $
��3.6* ��� $
��1.4* �.+�� $��0��$��10q�.+��� $��0��$��0q�$��48�����#28719e$��25(� �� 
IsReadOnly . $��Verdana)�$
��Bold"�
�������$	��Top=�.��PARSER� $��2��.+���  �� Orientation$ 
Horizontal=�$��Right=�.+��� $��10q�$��Resources/Table.png��$��30��$��Right=��� -    ��Name$ 
LoadButton��$��	10,0,20,0q�$��История��$��0q����#28719e$��15(�$
��Bold"�$��Verdana)�$��Center=�� $��1�����#d3dfe3 �� $
��1.5* ��� $��5* �.+�� $��0�� �� $��2* �$	��160���� $��9* �.+��� $��0��$��	20,20,0,0q�.+��� -   $ ParseButton��$��Обновить��$��7q�$��0q����#d3dfe3$��20(����#2f3233$
��Bold"�$��Verdana)�$
��Left=��� -   $ 
SaveButton��$��Сохранить��$��7q�$��0q����#d3dfe3$��20(����#2f3233$
��Bold"�$��Verdana)�$
��Left=��� -   $ ShortInfoButton��$��Просмотреть��$��7q�$��0q����#d3dfe3$��20(����#2f3233$
��Bold"�$��Verdana)�$
��Left=�� $��1�� �� $��48 ��� $��1* �.+��� $��0�����#28719e$��48�����#28719e$��25(� . $��Verdana)�$
��Bold"�
�������$	��Top=�$��Center=�$��Center=�.��Инфомация� $��1��$��0, 0, 0, 20q�.+���  �� TextWrapping$
 Wrap=� . $��14(�$��0q����#2f3233$
��Bold"�$��Verdana)�.��.����  �� Path$ CurrentNote�	 �� StringFormat$	 Запись: {0}��� $��1��$��20q����#d3dfe3 �� $��7* ��� $��* �.+�&   System.Windows.Controls.DataGrid  -   
  Name$
 	DataGrid1��$��0��   
IsReadOnly .    	RowHeight$ 40Q����#2f3233$
��Bold"�$��Verdana)�   CanUserAddRows .     CanUserDeleteRows .     CanUserReorderColumns .     CanUserResizeRows .     CanUserResizeColumns .  ����    Columns 0  *System.Windows.Controls.DataGridTextColumn  ,  &System.Windows.Controls.DataGridColumn   Header$- 'Идентификатор угрозы��   Width5  /System.Windows.Controls.DataGridLengthConverter$ * 1  +System.Windows.Controls.DataGridBoundColumn   Binding �� *IdInfo+  $+ %Наименование угрозы��$ 5*  �� *Name+� $��1��$	��Top=����#d3dfe3 �� $��30 ��� $��60 �.+��� $��0��$��Center=�$��14(�$��0q����#2f3233$
��Bold"�$��Verdana)����� $ CurrentPage�$	 Страница: {0}��� $��1�����#d3dfe3 �� $��* ��� $��* �.+��� -   $��0��$��
Назад��$��2q�$��20(����#2f3233$
��Bold"�$��Verdana)��� -   $��1��$��Вперед��$��2q�$��20(����#2f3233$
��Bold"�$��Verdana)�