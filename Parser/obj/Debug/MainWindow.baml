   M S B A M L   `   `   `  ����A  =Parser, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullclr-namespace:ParserParser  5   	       Parser.MainWindow   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   &localclr-namespace:Parser     5   	     � Title$   ��5   	   $	��900��6   $
��1600��6   $	��450��6,   $	��800��6<    � Icon$ Resources/iconParse.ico��5	   	   .��5
      �  � RowDefinitions 5   
   �� 5      $��50 �6   �� 5      $��8* �6   .+�5   
   �� $��0��6   ���#28719e6   � 5   
   $��0��6    � ColumnDefinitions 5      �� 5      $
��1.5* �6#   �� 5      $
��3.6* �6#   �� 5      $
��1.4* �6#   .+�5      � $��0��6   $��10q�6    .+�5      �� $��0��6   $��0q�6'   $��48��6;   ���#28719e$��25(�6\    �� 
IsReadOnly . $��Verdana)�6|   $
��Bold"�6�   
�������$	��Top=�6�   .��6�   PARSER5      5       � 5"      $��2��6   .+�5#      ��  �� Orientation$ 
Horizontal=�6   $��Right=�66   .+�5$      �� $��10q�6   $��Resources/Table.png��6(   $��30��6E   $��Right=�6P   �� 5%      -   5         ��Name$ 
LoadButton��5%   0   $��	10,0,20,0q�6   $��История��6`   $��0q�6r   ���#28719e$��15(�6�   $
��Bold"�6�   $��Verdana)�6�   $��Center=�6�   5'      5(      5)      � 5+   
   $��1��6   ���#d3dfe3 5,      �� 5-      $
��1.5* �6#   �� 5.      $��5* �6#   .+�51      � $��0��6    52      �� 53      $��2* �6$   $	��160��60   �� 54      $��9* �6$   .+�56      �� $��0��6   $��	20,20,0,0q�6*   .+�57      �� -   5        $ ParseButton��57      $��Обновить��6L   $��7q�6_   $��0q�6j   ���#d3dfe3$��20(�6�   ���#2f3233$
��Bold"�6�   $��Verdana)�6�   $
��Left=�6�   �� 58      -   5        $ 
SaveButton��58      $��Сохранить��6J   $��7q�6^   $��0q�6i   ���#d3dfe3$��20(�6�   ���#2f3233$
��Bold"�6�   $��Verdana)�6�   $
��Left=�6�   �� 59      -   5        $ ShortInfoButton��59      $��Просмотреть��6T   $��7q�6j   $��0q�6u   ���#d3dfe3$��20(�6�   ���#2f3233$
��Bold"�6�   $��Verdana)�6�   $
��Left=�6�   5:      � 5;      $��1��6    5<      �� 5=      $��48 �6(   �� 5>      $��1* �6(   .+�5@      �� $��0��6   ���#28719e$��48��6A   ���#28719e$��25(�6b    . $��Verdana)�6�   $
��Bold"�6�   
�������$	��Top=�6�   $��Center=�6�   $��Center=�6�   .��6  Инфомация5B      � 5C      $��1��6   $��0, 0, 0, 20q�6(   .+�5D      ��  �� TextWrapping$
 Wrap=�6"    . $��14(�6H   $��0q�6V   ���#2f3233$
��Bold"�6   $��Verdana)�6�   .��6�   .��5E      ��  �� Path$ CurrentNote�6&   	 �� StringFormat$	 Запись: {0}��5F   Y   5G      5H      5J      5L      5M      � 5O      $��1��6   $��20q�6#   ���#d3dfe3 5P      �� 5Q      $��7* �6$   �� 5R      $��* �6$   .+�5U      &   System.Windows.Controls.DataGrid  -   5        
  Name$
 	DataGrid1��5U      $��0��6F      
IsReadOnly .    	RowHeight$ 40Q�6�   ���#2f3233$
��Bold"�6�   $��Verdana)�6�      CanUserAddRows .     CanUserDeleteRows .     CanUserReorderColumns .     CanUserResizeRows .     CanUserResizeColumns .  ��6.   ��    Columns 5V      0  *System.Windows.Controls.DataGridTextColumn  5W      ,  &System.Windows.Controls.DataGridColumn   Header$- 'Идентификатор угрозы��6-      Width5  /System.Windows.Controls.DataGridLengthConverter$ * 6K   1  +System.Windows.Controls.DataGridBoundColumn   Binding 6U   �� *IdInfo+  5X      $+ %Наименование угрозы��6-   $ 5* 6J    6U   �� *Name+5Z      � 5[      $��1��6   $	��Top=�6$   ���#d3dfe3 5\      �� 5]      $��30 �6(   �� 5^      $��60 �6(   .+�5`      �� $��0��6   $��Center=�6)   $��14(�6F   $��0q�6T   ���#2f3233$
��Bold"�6}   $��Verdana)�6�   ��5a      �� 5b      $ CurrentPage�6&   $	 Страница: {0}��5c   [   5d      5f      � 5g      $��1��6   ���#d3dfe3 5h      �� 5i      $��* �6/   �� 5j      $��* �6/   .+�5l      �� -   5        $��0��5l   !   $��
Назад��61   $��2q�6[   $��20(�6f   ���#2f3233$
��Bold"�6�   $��Verdana)�6�   �� 5m      -   5        $��1��5m   !   $��Вперед��61   $��2q�6X   $��20(�6c   ���#2f3233$
��Bold"�6�   $��Verdana)�6�   5n      5p      5r      5s      5u      5v      