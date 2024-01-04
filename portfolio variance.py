# -*- coding: utf-8 -*-
"""
Created on Fri Apr 28 14:13:41 2023


"""
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk


class PortfolioGenerator:
    def __init__(self, file_path):
        self.file_path = file_path
        self.df = None
        self.selected_tickers = None
        self.portfolio = None
        self.returns = None
        self.covariance = None

    def read_data(self):
        self.df = pd.read_excel(self.file_path, index_col=0)
        self.df = self.df.pivot_table(index='date', columns='Name', values='close')
        self.df.dropna(inplace=True)
        self.df.sort_index(inplace=True)

    def select_tickers(self, num_tickers):
        tickers = self.df.columns.values
        self.selected_tickers = np.random.choice(tickers, num_tickers, replace=False)


    def create_portfolio(self):
        self.portfolio = self.df[self.selected_tickers].copy()
        return self.portfolio

    def save_portfolio(self, sheet_name='Portfolio'):
        if self.portfolio is None:
            raise ValueError("Portfolio not created yet")

    def calculate_returns(self, sheet_name='Returns'):
        if self.portfolio is None:
            raise ValueError("Portfolio not created yet")

        self.returns = self.portfolio.pct_change(periods=4)
        self.returns = self.returns.dropna()
        return self.returns
    
    def calculate_covariance(self, sheet_name='Covariance'):
        if self.returns is None:
            raise ValueError("Returns not calculated yet")

        self.covariance = self.returns.cov()
        return self.covariance
    
    def calculate_portfolio_variance(self, max_num_assets):
        if self.portfolio is None:
            raise ValueError("Portfolio not created yet")
        if self.covariance is None:
            raise ValueError("Covariance matrix not calculated yet")

        
        variances = []
        for num_assets in range(1, max_num_assets+1):
             weights = np.ones(num_assets) / num_assets
             selected_tickers = self.portfolio.columns.values[:num_assets]
             covariance = self.covariance.loc[selected_tickers, selected_tickers]
             sub_variance = np.dot(weights.T, np.dot(covariance, weights))
             variances.append(sub_variance)
        return variances


class PortfolioGeneratorGUI:
    def __init__(self, master):
        self.master = master
        self.master.title('Portfolio Generator')
        
        self.frame1 = ttk.Frame(self.master)
        self.frame1.pack(pady=10)
        self.start_button = ttk.Button(self.frame1, text='Réaliser le test', command=self.show_page2)
        self.start_button.pack(padx=10, pady=10)
        
        self.frame2 = ttk.Frame(self.master)
        self.num_assets_label = ttk.Label(self.frame2, text='Quel est le nombre de titres ?')
        self.num_assets_label.pack(padx=10, pady=10)
        self.num_assets_entry = ttk.Entry(self.frame2, width=10)
        self.num_assets_entry.pack(padx=10, pady=5)
        self.validate_button = ttk.Button(self.frame2, text='Valider', command=self.show_page3)
        self.validate_button.pack(padx=10, pady=10)
        
        self.frame3 = ttk.Frame(self.master)
        self.close_button = ttk.Button(self.frame3, text='Fermer', command=self.close)
        self.close_button.pack(padx=10, pady=10)
        self.download_button = ttk.Button(self.frame3, text='Télécharger le graphique', command=self.download)
        self.download_button.pack(padx=10, pady=10)
        self.figure_canvas = None
        
    def show_page2(self):
        self.frame1.pack_forget()
        self.frame2.pack(pady=10)
        
    def show_page3(self):
        try:
            num_assets = int(self.num_assets_entry.get())
            generator = PortfolioGenerator('NYSE_2015_to_2016.xlsx')
            generator.read_data()
            generator.select_tickers(num_assets)
            portfolio = generator.create_portfolio()
            generator.calculate_returns()
            generator.calculate_covariance()
            variances = generator.calculate_portfolio_variance(num_assets)
            self.plot_portfolio_variance(variances)
            self.frame2.pack_forget()
            self.frame3.pack(pady=10)
        except ValueError:
            pass
        
    def plot_portfolio_variance(self, variances):
        if self.figure_canvas is not None:
            self.figure_canvas.get_tk_widget().destroy()
        
        fig, ax = plt.subplots(figsize=(8, 5))
        num_assets_range = range(1, len(variances)+1)
        ax.plot(num_assets_range, variances)
        ax.set_xlabel('Nombre de titres')
        ax.set_ylabel('Variance du portefeuille')
        ax.set_title('Variance du portefeuille en fonction du nombre de titres')
        self.figure_canvas = FigureCanvasTkAgg(fig, master=self.frame3)
        self.figure_canvas.get_tk_widget().pack(pady=10)
        
    def download(self):
        filetypes = [('PNG', '.png'), ('PDF', '.pdf'), ('SVG', '.svg')]
        filepath = filedialog.asksaveasfilename(filetypes=filetypes, defaultextension=filetypes)
        if filepath:
            self.figure_canvas.figure.savefig(filepath) 
   
    def close(self):
        self.master.destroy()

root = tk.Tk()
app = PortfolioGeneratorGUI(root)
root.mainloop()



