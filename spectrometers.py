import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
from IPython.display import display
from randomcolor import RandomColor

tables_directory = './'

df = pd.read_excel(tables_directory + 'iPHAOS.xlsx',
                   sheet_name='spectrometers', dtype = None, header=0, skiprows=1)
# to load all the sheets use None in the sheet_name
# df = pd.read_excel(tables_directory + 'iPHAOS.xlsx',
#                   sheet_name=None, header=0, skiprows=1)

# to print to screen the imported dataframe
# display(df)

fig = plt.figure(figsize=plt.figaspect(0.5))


df['Bandwidth-to-resolution ratio'] = pd.to_numeric(df['Bandwidth-to-resolution ratio'], errors='coerce')

# example 1 shown in the paper (Fig. 10)
df.plot(x='Ref.', y='Bandwidth [nm]', style='o')

# example 2 shown in the paper (Fig. 11)
df.plot(x='Ref.', y='Spectral resolution [nm]', style='o')

# example 3 shown in the paper (Fig. 12)
df_dropped1 = df.dropna(subset=['Bandwidth [nm]', 'Spectral resolution [nm]','Bandwidth-to-resolution ratio','Ref.'])
ax = fig.add_subplot(projection='3d')
# create a RandomColor object
rc = RandomColor()
# generate random colors and markers, as many as the lines corresponding to the intersection of column 'Spectral resolution [nm]' and column 'Bandwidth-to-resolution ratio'
num_points = len(df_dropped1)
colors = rc.generate(count = num_points)
markers = Line2D.filled_markers
# plot each point with a unique color and shape
for i, (x, y, z, reference) in enumerate(zip(df_dropped1['Spectral resolution [nm]'], df_dropped1['Bandwidth-to-resolution ratio'], df_dropped1['Bandwidth [nm]'], df_dropped1['Ref.'])):
    ax.scatter(x, y, z, color=colors[i], marker=markers[i % len(markers)], label=reference)
# set labels and title
ax.set_xlabel('Spectral resolution [nm]')
ax.set_ylabel('Bandwidth-to-resolution ratio')
ax.set_zlabel('Bandwidth [nm]')
# ax.set_title("Integrated Spectrometers Metrics")

ax.set_xlim(0, 20)
ax.set_ylim(0, 800)
ax.set_zlim(0, 2000)


# show legend
ax.legend(bbox_to_anchor=(1.05, 1.0), loc='upper left', title = 'Ref.')
plt.show()

print('done!')

