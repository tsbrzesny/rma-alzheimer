   M S B A M L   `   `   `  ����  	WPF_Roots    WPF_Roots.PromptBox   5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   -   5          � Title$  	PromptBox��5   	    � WindowStartupLocation$ CenterOwner=�6    � SizeToContent$ WidthAndHeight=�5   	    � 
ResizeMode$ NoResize=�6(    � WindowStyle$
 None=�6>    J� Effect 5      3  -System.Windows.Media.Effects.DropShadowEffect  5	   
   .��5      �� 
�������$��1q�6    ._�5   
   ]� 
�������.+�5      �� -   5        $��tb_title��5      $	��Top=�6(   $��16(�6=    �� Padding$ 5,5,5,5q�6K   ���	#FF4040FF
�������Title5      ]� 5      .+�5      �� -   5        $��i_icon��5      $��Images/prompt.png��6&   $
��Left=�6A   $��70��6W   $��70��6c   $��
15,15,20,0q�5      $	��Top=�6,    �� BitmapScalingMode$ HighQuality=�6D   �� 5      .+�5      �� -   5        $��tb_msg��5       $��	0,15,15,0q�6.   	Message6J   �� 5      -   5        $��cc_msg��5   %   $��
0,15,15,15q�63   ]� 5      -   5        $��frame_buttons��5       $��Bottom=�65   .+�5      �� -   5        $��bt_Ok��5   !   $��Right=�6.   $	��100��6E   $��
0,15,15,15q�6Q    �� 	IsDefault . .��6�   Ok6�   �� 5      -   5        $��	bt_Cancel��5   !   $��Right=�62   $	��100��6I   $��
0,15,15,15q�6U   	 �� IsCancel	 . .��6y   Cancel6�   � 5      5      5       5!      5"      5#      5$      