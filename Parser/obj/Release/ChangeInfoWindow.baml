   M S B A M L   `   `   `  ���� A  =Parser, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullclr-namespace:ParserParser      Parser.ChangeInfoWindow   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     &localclr-namespace:Parser       � Title$   ��$	��900��$
��1600��$	��450��$	��800�� � Icon$ Resources/iconParse.ico��.���  � RowDefinitions �� $��50 ��� $��* ��� $��50 �.+�� $��0�� � ColumnDefinitions �� $��4* ��� $��1* �.+��� $��0�� �� Orientation$ 
Horizontal=�$��Center=�.+���  �� 
IsReadOnly . $��Center=�$��14(�$��0q����#2f3233$
��Bold"�$��Verdana)�.��:8Количество измененных записей�� -    ��Name$ changedCount�� . $��Center=�$��14(�$��0q����#2f3233$
��Bold"�$��Verdana)�.��0�� -   $ statusTextBox��$��1�� . $��Center=�$��Center=�$��0q�$��20(�$��Verdana)�.��status� $��1�� �� $��1* ��� $��1* �.+��� -    ��Name$ 
lvSheetOld��$��0�� �� View �� 	 �� AllowsColumnReorder	 .  .P��� 
 �� Width$
 30�� �� DisplayMemberBinding ��  �� Path$ Id�.��ID�� $	
 100�� �� $
 Name�.��Наименование�� $	
 100�� �� $ Description�.��Описание�� $	
 100�� �� $ Source�.��Источник�� $	
 100�� �� $ Threat�.��%#Объект воздействия�� $
 80�� �� $ IsNotConfidential�.��97Нарушение конфиденциальности�� $
 80�� �� $ 
IsComplete�.��+)Нарушение целостности�� $
 80�� �� $ IsAccessible�.��#!Нарушение доступа�� $
 60�� �� $ DateIn�.��Дата создания�� $
 60�� �� $ DateRewrite�.��Дата изменения�� -   $ 
lvSheetNew��$��1�� �� 	 .  .P��� $
 30�� �� $ Id�.��ID�� $	
 100�� �� $
 Name�.��Наименование�� $	
 100�� �� $ Description�.��Описание�� $	
 100�� �� $ Source�.��Источник�� $	
 100�� �� $ Threat�.��%#Объект воздействия�� $
 80�� �� $ IsNotConfidential�.��97Нарушение конфиденциальности�� $
 80�� �� $ 
IsComplete�.��+)Нарушение целостности�� $
 80�� �� $ IsAccessible�.��#!Нарушение доступа�� $
 60�� �� $ DateIn�.��Дата создания�� $
 60�� �� $ DateRewrite�.��Дата изменения� $��2��.+��� -    ��Name$ OkButton��$
��Ок��$��Center=�$��0q�$��20(����#2f3233$
��Bold"�$��Verdana)�