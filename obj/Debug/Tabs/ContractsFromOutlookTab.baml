   M S B A M L   `   `   `  ����I  EShellBeeHelper, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null9!clr-namespace:ShellBeeHelper.TabsShellBeeHelper.Tabs  5      1    +ShellBeeHelper.Tabs.ContractsFromOutlookTab   5      W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5      R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5      Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5      @d2http://schemas.microsoft.com/expression/blend/2008     5      3local!clr-namespace:ShellBeeHelper.Tabs     5      .��5	      �   � RowDefinitions  5
   
   �� 5      $
��auto �6   �� 5      $
��auto �6   �� 5      $
��auto �6   �� 5      $
��auto �6    � ColumnDefinitions 5   
   �� 5      $
��auto �6   �� 5      $��* �6   .+�5   
   �� -   5         ��Name$ EmailAddressLabel��5      $��0��6+   $��0��68   $��Right=�6H   $��5q�6d   $��Email Address: ��5      �� 5   
   -   5         ��Name$ EmailAddressTextBox��5      $��0��6/   $��1��6<   $��Stretch=�6L   $��5q�6j   5      �� 5   
   -   5        $ SourceFolderLabel��5      $��1��6+   $��0��68   $��Right=�6H   $��5q�6d   $��Source Folder: ��5      �� 5   
   -   5        $ SourceFolderTextBox��5      $��1��6/   $��1��6<   $��Stretch=�6L   $��5q�6j   �� 5   
   -   5        $ DestFolderLabel��5      $��2��6)   $��0��66   $��Right=�6F   $��5q�6b   $��Destination Folder: ��5      �� 5    
   -   5        $ DestFolderTextBox��5       $��2��6-   $��1��6:   $��Stretch=�6J   $��5q�6h   �� 5"   
   -   5         ��Name$ 
ScanButton��5"      $��3��6%   $��2��62   $��Center=�6F   $��5q�6c   $��Scan for Contracts��5#      $��5q�6.   5$      5%      5&      