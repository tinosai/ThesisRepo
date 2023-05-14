convert fig1.png fig2.png +append fig12.png
convert fig3.png fig4.png +append fig34.png

convert fig12.png fig34.png -append fig_complete.png
