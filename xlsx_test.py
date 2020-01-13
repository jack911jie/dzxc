import matplotlib.pyplot as plt
import numpy as np
import mpl_toolkits.axisartist as axisartist

# pi=np.pi
# print(np.cos(pi))

fig=plt.figure()
#ax=fig.add_subplot(111)
ax=axisartist.Subplot(fig,111)
fig.add_axes(ax)
ax.axis[:].set_visible(False)
ax.axis['x']=ax.new_floating_axis(0,0)
ax.axis['x'].set_axisline_style('->',size=1)
ax.axis["y"] = ax.new_floating_axis(1,0)
ax.axis["y"].set_axisline_style("-|>", size = 1.0)
#ax.set(xlim=[0.5,4.5],ylim=[-2,8],title='test',ylabel='y',xlabel='x')



plt.show()