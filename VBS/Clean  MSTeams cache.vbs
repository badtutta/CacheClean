 
 D i m   o b j S h e l l  
 S e t   o b j S h e l l = C r e a t e O b j e c t ( " W S c r i p t . S h e l l " )  
  
 ' e n t e r   t h e   P o w e r S h e l l   e x p r e s s i o n  
 s t r E x p r e s s i o n = " . \ C a c h e C l e a n . p s 1   - M S T e a m s "  
  
 s t r C M D = " C : \ W i n d o w s \ S y s t e m 3 2 \ W i n d o w s P o w e r S h e l l \ v 1 . 0 \ p o w e r s h e l l . e x e   - n o l o g o   - c o m m a n d   "   &   C h r ( 3 4 )   & _  
 " & { "   &   s t r E x p r e s s i o n   & " } "   &   C h r ( 3 4 )  
  
 ' U n c o m m e n t   n e x t   l i n e   f o r   d e b u g g i n g  
 ' W S c r i p t . E c h o   s t r C M D  
  
 ' u s e   0   t o   h i d e   w i n d o w  
 o b j S h e l l . R u n   s t r C M D , 0 