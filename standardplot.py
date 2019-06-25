import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
import numpy as np

# TODO: Studienverlauf graphisch darstellen
# @Return: Ausgabe des Studienverlaufs mittels ECTS-Punkte

x = [30, 60, 90, 120, 150, 180, 210]
y = [25, 50, 80, 115, 145, 170, 210]

np.arange(len(x))

plt.plot(x, y)
plt.xticks(x, (x))
plt.title('Verlauf der bestanden Prüfungen')
plt.xlabel('Soll-ECTS')
plt.ylabel('Ist-ECTS')
plt.show()

# TODO: Graphische Darstellung der weltweit größten Autombilzulieferer
# @Return: Ausgabe der Zuliefer nach Umsatz

x = np.arange(4)
money = [33.5e6, 36.4e6, 44.0e6, 48.5e6]


def billions(x, pos):
    return '$%1.1fMrd' % (x * 1e-6)


formatter = FuncFormatter(billions)

fig, ax = plt.subplots()
ax.yaxis.set_major_formatter(formatter)
plt.title('Größten Automobilzulieferer weltweit')
plt.bar(x, money, color="limegreen")
plt.xticks(x, ('ZF', 'Denso', 'Continental', 'Bosch'))
plt.show()
