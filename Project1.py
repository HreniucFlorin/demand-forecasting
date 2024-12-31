import pandas as pd
import locale
from prophet import Prophet
import plotly.graph_objects as go
import plotly.io as pio
from tkinter import Tk, filedialog, Button, Label, messagebox, Text, END, Toplevel, Entry, Frame

# Setăm browserul implicit pentru Plotly
pio.renderers.default = "browser"

# Variabile globale pentru date și prognoză
plotly_data = None
plotly_forecast = None

# Funcție pentru citirea datelor și generarea prognozei
def load_and_forecast():
    global plotly_data, plotly_forecast

    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if not file_path:
        messagebox.showinfo("Informație", "Niciun fișier selectat!")
        return

    try:
        data = pd.read_excel(file_path, sheet_name="Prognoza_Cerere_Laptopuri")
        data.columns = ['Luna', 'Cerere', 'Preț', 'Venituri', 'Cerere Cumulată', 'Rata de Creștere', 'Prognoză Cerere']
        data['ds'] = pd.date_range(start='2025-01-01', periods=len(data), freq='M')
        data['y'] = data['Cerere']

        model = Prophet()
        model.fit(data[['ds', 'y']])

        future = model.make_future_dataframe(periods=12, freq='M')
        forecast = model.predict(future)

        avg_demand = int(data['y'].mean())
        max_demand = int(forecast['yhat'].max())
        min_demand = int(forecast['yhat'].min())
        total_demand_next_year = int(forecast['yhat'].iloc[len(data):].sum())
        next_month_forecast = int(forecast['yhat'].iloc[len(data)])
        next_month_revenue = int(next_month_forecast * data['Preț'].iloc[-1])
        locale.setlocale(locale.LC_TIME, 'Romanian_Romania')
        max_month = forecast.loc[forecast['yhat'].idxmax(), 'ds'].strftime('%B %Y').capitalize()
        min_month = forecast.loc[forecast['yhat'].idxmin(), 'ds'].strftime('%B %Y').capitalize()

        recommendations = "🔍 ** Recomandări **\n\n"
        if next_month_forecast < data['y'].iloc[-1]:
            recommendations += "🔽 Se recomandă o reducere a prețului produsului cu 20% pentru a stimula cererea în luna următoare.\n"
        else:
            recommendations += "🔼 Creșterea cererii sugerează lansarea unei campanii promoționale pentru a maximiza veniturile.\n"

        if forecast['yhat'].diff().max() > 10:
            recommendations += "📈 Creșterea prognozată a cererii sugerează optimizarea stocurilor pentru următoarele luni.\n"

        recommendations += "🏭 Se recomandă ajustarea planificării producției pentru a răspunde cererii din luna Ianurie 2025.\n"
        recommendations += "🛒 Luați în considerare extinderea canalelor de distribuție pentru a evita pierderile din cauza lipsei de stoc.\n"

        calculations = (
            "📊 ** Calcule Detaliate **\n\n"
            f"• Media cererii istorice: {avg_demand} unități\n"
            f"• Prognoza cererii maxime: {max_demand} unități (în luna {max_month})\n"
            f"• Prognoza cererii minime: {min_demand} unități (în luna {min_month})\n"
            f"• Total cerere prognozată pentru anul următor: {total_demand_next_year} unități\n"
            f"• Venituri estimate pentru luna următoare: {next_month_revenue} unități monetare\n"
        )

        predictions = (
            "🔮 ** Predicții și Observații **\n\n"
            f"• Cererea este {'🔼 în creștere' if next_month_forecast > data['y'].iloc[-1] else '🔽 în scădere'} "
            f"cu {abs(next_month_forecast - data['y'].iloc[-1]):.0f} unități față de luna anterioară\n"
            f"• Cea mai mare creștere lunară prognozată este de {forecast['yhat'].diff().max():.0f} unități\n"
            f"• Cea mai mare scădere lunară prognozată este de {forecast['yhat'].diff().min():.0f} unități\n"
            f"• Luna {max_month} este prognozată să aibă cea mai mare cerere\n"
            f"• Luna {min_month} este prognozată să aibă cea mai mică cerere\n"
        )

        stats_text.delete(1.0, END)
        stats_text.insert(END, calculations)
        stats_text.insert(END, predictions)
        stats_text.insert(END, recommendations)

        plotly_data = data
        plotly_forecast = forecast

    except Exception as e:
        messagebox.showerror("Eroare", f"Eroare la citirea datelor sau generarea prognozei: {e}")


def forecast_with_plotly(data, forecast):
    if data is None or forecast is None:
        messagebox.showerror("Eroare", "Nu există date disponibile pentru a genera graficul. Încărcați un fișier mai întâi!")
        return

    # Diversificăm prognozele pentru viitor
    import numpy as np

    # Aplicați fluctuații mai mari asupra valorilor prognozate
    fluctuation_factor = 0.1  # 10% fluctuație
    forecast['yhat'] = forecast['yhat'] * (1 + np.random.uniform(-fluctuation_factor, fluctuation_factor, len(forecast)))

    # Calculăm veniturile prognozate pe baza fluctuațiilor
    forecast['Venituri'] = forecast['yhat'] * data['Preț'].iloc[-1]

    # Calculăm rata de creștere pentru datele viitoare
    forecast['Rata de Creștere'] = forecast['yhat'].pct_change() * 100

    # Funcție internă pentru generare de grafic specific
    def generate_graph(graph_type):
        fig = go.Figure()

        # Date combinate pentru prognoză
        combined_ds = pd.concat([data['ds'], forecast['ds'][len(data):]])
        combined_cerere = pd.concat([data['y'], forecast['yhat'][len(data):]])
        combined_rata = pd.concat([data['Rata de Creștere'].fillna(0), forecast['Rata de Creștere'][len(data):]])
        combined_cumulativa = combined_cerere.cumsum()
        combined_venituri = pd.concat([data['Venituri'], forecast['Venituri'][len(data):]])

        # Configurare grafic în funcție de selecție
        if graph_type == "Cerere":
            fig.add_trace(go.Scatter(
                x=combined_ds, y=combined_cerere,
                mode='lines+markers', name='Cerere',
                line=dict(color='blue', width=2),
                marker=dict(size=8, color=['blue'] * len(data) + ['orange'] * (len(forecast) - len(data)))
            ))

        elif graph_type == "Preț":
            fig.add_trace(go.Scatter(
                x=combined_ds, y=pd.concat([data['Preț'], forecast['yhat'][len(data):]]),
                mode='lines+markers', name='Preț',
                line=dict(color='green', width=2),
                marker=dict(size=8, color=['green'] * len(data) + ['orange'] * (len(forecast) - len(data)))
            ))

        elif graph_type == "Venituri":
            fig.add_trace(go.Scatter(
                x=combined_ds, y=combined_venituri,
                mode='lines+markers', name='Venituri',
                line=dict(color='purple', width=2),
                marker=dict(size=8, color=['purple'] * len(data) + ['orange'] * (len(forecast) - len(data)))
            ))

        elif graph_type == "Rata de Creștere":
            fig.add_trace(go.Scatter(
                x=combined_ds, y=combined_rata,
                mode='lines+markers', name='Rata de Creștere',
                line=dict(color='orange', width=2),
                marker=dict(size=8, color=['orange'] * len(data) + ['red'] * (len(forecast) - len(data)))
            ))

        elif graph_type == "Cerere Cumulativă":
            fig.add_trace(go.Scatter(
                x=combined_ds, y=combined_cumulativa,
                mode='lines+markers', name='Cerere Cumulativă',
                line=dict(color='brown', width=2),
                marker=dict(size=8, color=['brown'] * len(data) + ['orange'] * (len(forecast) - len(data)))
            ))

        # Configurare generală grafic
        fig.update_layout(
            title=f"Grafic: {graph_type}",
            xaxis_title="Data",
            yaxis_title="Valoare",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        )
        fig.show()

    # Afișează meniul pentru selecție
    def show_selection():
        selection_window = Toplevel()
        selection_window.title("Selectați graficul")
        selection_window.geometry("300x300")

        Button(selection_window, text="Cerere", command=lambda: [generate_graph("Cerere"), selection_window.destroy()]).pack(pady=10)
        Button(selection_window, text="Preț", command=lambda: [generate_graph("Preț"), selection_window.destroy()]).pack(pady=10)
        Button(selection_window, text="Venituri", command=lambda: [generate_graph("Venituri"), selection_window.destroy()]).pack(pady=10)
        Button(selection_window, text="Rata de Creștere", command=lambda: [generate_graph("Rata de Creștere"), selection_window.destroy()]).pack(pady=10)
        Button(selection_window, text="Cerere Cumulativă", command=lambda: [generate_graph("Cerere Cumulativă"), selection_window.destroy()]).pack(pady=10)

    show_selection()

# Funcție pentru simularea impactului unui preț nou
def simulate_price_impact():
    def calculate_impact():
        try:
            new_price = float(price_entry.get())
            current_price = plotly_data['Preț'].iloc[-1]
            current_demand = plotly_data['y'].iloc[-1]
            demand_change = (new_price - current_price) / current_price
            new_demand = int(current_demand * (1 - demand_change))
            messagebox.showinfo("Rezultat Simulare", f"Cererea estimată: {new_demand} unități.")
            simulation_window.destroy()  # Închide fereastra după calcul
        except ValueError:
            messagebox.showerror("Eroare", "Introduceți un preț valid!")

    simulation_window = Toplevel()
    simulation_window.title("Simulare Impact Preț")
    simulation_window.geometry("300x150")
    Label(simulation_window, text="Introduceți un preț nou:", font=("Arial", 12)).pack(pady=10)
    price_entry = Entry(simulation_window, font=("Arial", 12))
    price_entry.pack(pady=5)
    Button(simulation_window, text="Calculează", command=calculate_impact, font=("Arial", 12), bg="#4682b4", fg="white").pack(pady=10)

# Funcție pentru simularea impactului profitului
def simulate_profit_impact():
    def calculate_profit():
        try:
            new_price = float(price_entry.get())
            cost_per_unit = float(cost_entry.get())
            current_price = plotly_data['Preț'].iloc[-1]
            current_demand = plotly_data['y'].iloc[-1]
            demand_change = (new_price - current_price) / current_price
            new_demand = int(current_demand * (1 - demand_change))
            profit = (new_price - cost_per_unit) * new_demand
            messagebox.showinfo("Rezultat Simulare", f"Profitul estimat: {profit:.0f} unități monetare.")
            simulation_window.destroy()  # Închide fereastra după calcul
        except ValueError:
            messagebox.showerror("Eroare", "Introduceți valori valide pentru preț și cost!")

    simulation_window = Toplevel()
    simulation_window.title("Simulare Impact Profit")
    simulation_window.geometry("300x200")
    Label(simulation_window, text="Introduceți un preț nou:", font=("Arial", 12)).pack(pady=5)
    price_entry = Entry(simulation_window, font=("Arial", 12))
    price_entry.pack(pady=5)
    Label(simulation_window, text="Introduceți costul pe unitate:", font=("Arial", 12)).pack(pady=5)
    cost_entry = Entry(simulation_window, font=("Arial", 12))
    cost_entry.pack(pady=5)
    Button(simulation_window, text="Calculează", command=calculate_profit, font=("Arial", 12), bg="#4682b4",
           fg="white").pack(pady=10)

# Creează interfața
def create_gui():
    global stats_text, plotly_data, plotly_forecast

    root = Tk()
    root.title("Sistem de prognoză cerere")
    root.geometry("900x750")
    root.configure(bg="#1e3d59")

    Label(root, text="Sistem de Prognoză a Cererii", font=("Arial", 22, "bold"), fg="#ffffff", bg="#1e3d59").pack(
        pady=20)

    Label(root,
          text="Această aplicație vă permite să încărcați date istorice despre cererea produselor și să generați prognoze detaliate.",
          font=("Arial", 12), bg="#1e3d59", fg="#cce7e8", wraplength=800, justify="center").pack(pady=10)

    Label(root, text="Funcții disponibile: încărcare fișier Excel, generare prognoză, vizualizare grafic interactiv.",
          font=("Arial", 10, "italic"), bg="#1e3d59", fg="#cce7e8").pack(pady=5)

    Button(root, text="Încarcă fișier și calculează", command=load_and_forecast, font=("Arial", 12), bg="#4682b4",
           fg="#ffffff", width=25).pack(pady=10)

    Button(root, text="Grafic Interactiv", command=lambda: forecast_with_plotly(plotly_data, plotly_forecast),
           font=("Arial", 12), bg="#4682b4", fg="#ffffff", width=25).pack(pady=10)

    # Creează un container pentru simulările de preț și profit
    simulation_frame = Frame(root, bg="#1e3d59")
    simulation_frame.pack(pady=10)

    # Buton Simulare Impact Preț
    Button(simulation_frame, text="Simulare Impact Preț", command=simulate_price_impact,
           font=("Arial", 12), bg="#4682b4", fg="#ffffff", width=25).grid(row=0, column=0, padx=5)

    # Buton Simulare Impact Profit
    Button(simulation_frame, text="Simulare Impact Profit", command=simulate_profit_impact,
           font=("Arial", 12), bg="#4682b4", fg="#ffffff", width=25).grid(row=0, column=1, padx=5)

    stats_text = Text(root, height=25, width=100, font=("Courier New", 12), wrap="word", bg="#2e4756",
                      fg="#ffffff")  # Casetă text albastru petrol închis
    stats_text.pack(pady=15)

    root.mainloop()


if __name__ == "__main__":
    create_gui()
