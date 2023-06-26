# myDAQ_app.py
"""Aplicación Python para medir componentes con myDAQ"""
"""Universidad de Salamanca - Raúl Rengel Estévez"""
"""Versión 2.0"""

import csv
import tkinter as tk
import nidaqmx.system
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from tkinter import ttk
from tkinter import scrolledtext as st
from tkinter.filedialog import asksaveasfilename
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg, NavigationToolbar2Tk)
import threading
from datetime import datetime

class BoundText(tk.Text):
    """Un widget de texto junto con una variable ligada"""
    
    def __init__(self, *args, textvariable=None, **kwargs):
        super().__init__(*args, **kwargs)
        self._variable = textvariable
        if self._variable:
            self.insert('1.0', self._variable.get())
            self._variable.trace_add('write', self._set_content)
            self.bind('<<Modified>>', self._set_var)
   
    def _set_content(self, *_):
        """Asocia los contenidos de texto a la variable"""
        self.delete('1.0', tk.END)
        self.insert('1.0', self._variable.get())
    
    def _set_var(self, *_):
        """Fija la variable a los contenidos de texto"""
        if self.edit_modified():
            content = self.get('1.0', 'end-1chars')
            self._variable.set(content)
            self.edit_modified(False)
            
class LabelInput(tk.Frame):
    """Widget que contiene una etiqueta y una entrada juntas"""
    def __init__(
            self, parent, label, var, input_class=ttk.Entry,
            input_args=None, label_args=None, **kwargs
            ):
        super().__init__(parent, **kwargs)
        input_args = input_args or {}
        label_args = label_args or {}
        self.input_class = input_class
        self.variable = var
        self.variable.label_widget = self
        if input_class in (ttk.Checkbutton, ttk.Button):
            input_args["text"] = label
        else:
            self.label = ttk.Label(self, text=label, **label_args)
            self.label.grid(row=0, column=0, sticky=(tk.W + tk.E))

        if input_class in (
                ttk.Checkbutton, ttk.Button, ttk.Radiobutton
                ):
            input_args["variable"] = self.variable
        else:
            input_args["textvariable"] = self.variable

        if input_class == ttk.Radiobutton:
            self.input = tk.Frame(self)
            for v in input_args.pop('values', []):
                button = ttk.Radiobutton(
                self.input, value=v, text=v, **input_args
                )
                button.pack(
                    side=tk.LEFT, ipadx=10, ipady=2, expand=True, fill='x'
                    )
        else:
            self.input = input_class(self, **input_args)
            
        self.input.grid(row=1, column=0, sticky=(tk.W + tk.E))
        self.columnconfigure(0, weight=1)
    
    def grid(self, sticky=(tk.E + tk.W), **kwargs):
        """Ignorar grid para añadir los valores sticky por defecto"""
        super().grid(sticky=sticky, **kwargs)

        
class DataRecordForm(ttk.Frame):
    """Formulario de entrada de los widgets"""
    def _add_frame(self, label, cols=3):
        """Añadir un LabelFrame al fromulario"""
        frame = ttk.LabelFrame(self, text=label)
        frame.grid(sticky=tk.W + tk.E)
        for i in range(cols):
            frame.columnconfigure(i, weight=1)
        return frame
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._vars = {
            'Tipo de medida': tk.StringVar(),
            'Ref': tk.StringVar(),
            'VDD Min': tk.DoubleVar(),
            'VDD Max': tk.DoubleVar(),
            'Incremento': tk.DoubleVar(),
            'VGS Min': tk.DoubleVar(),
            'VGS Max': tk.DoubleVar(),
            'IncrementoVGS': tk.DoubleVar(),
            'Valor de R (Ohm)': tk.DoubleVar(),
            'Consola': tk.StringVar()
        }

        self._vars["VDD Min"].set(-2)
        self._vars["VDD Max"].set(2)
        self._vars["Incremento"].set(0.02)
        self._vars["Valor de R (Ohm)"].set(100)

        self._vars["VGS Min"].set(0)
        self._vars["VGS Max"].set(5)
        self._vars["IncrementoVGS"].set(0.5)
        
        t_select = self._add_frame("Tipo de medida")
        
        LabelInput(
            t_select, "Selección de la medida a realizar", input_class=ttk.Radiobutton,
            var=self._vars['Tipo de medida'],
            input_args={"values": ["I-V Diodo", "Id-Vds MOS", "Id-Vgs MOS", "Ic-Vce BJT"]}
            ).grid(row=0, column=0)
             
        LabelInput(
            t_select, "Referencia del dispositivo", var=self._vars['Ref']
            ).grid(row=1, column=0)
        
        p_select = self._add_frame("Selección de parámetros")
        
        self.vddmin = LabelInput(
            p_select, "VDD Mínimo",
            input_class=ttk.Spinbox, var=self._vars['VDD Min'],
            input_args={"from_": -10, "to": 0, "increment": .01}
            )
        
        self.vddmax = LabelInput(
            p_select, "VDD Máximo",
            input_class=ttk.Spinbox, var=self._vars['VDD Max'],
            input_args={"from_": 0, "to": 10, "increment": .01}
            )
        
        self.incremento = LabelInput(
            p_select, "Incremento",
            input_class=ttk.Spinbox, var=self._vars['Incremento'],
            input_args={"from_": 0.01, "to": 1, "increment": .01}
            )
        
        self.vddmin.grid(row=0, column=0)
        self.vddmax.grid(row=0, column=1)
        self.incremento.grid(row=0, column=2)
        
        self.vgsmin = LabelInput(
            p_select, "VGS Mínimo",
            input_class=ttk.Spinbox, var=self._vars['VGS Min'],
            input_args={"from_": -1, "to": 5, "increment": .01}
            )
        
        self.vgsmax = LabelInput(
            p_select, "VGS Máximo",
            input_class=ttk.Spinbox, var=self._vars['VGS Max'],
            input_args={"from_": 0, "to": 10, "increment": .01}
            )
                      
        self.incrementovgs = LabelInput(
            p_select, "Incremento",
            input_class=ttk.Spinbox, var=self._vars['IncrementoVGS'],
            input_args={"from_": 0.01, "to": 1, "increment": .01}
            )
        
        self.vgsmin.grid(row=1, column=0)
        self.vgsmax.grid(row=1, column=1)
        self.incrementovgs.grid(row=1, column=2)
              
        self.resistencia = LabelInput(
            p_select, "Valor de R (Ohm)",
            input_class=ttk.Spinbox, var=self._vars['Valor de R (Ohm)'],
            input_args={"from_": 1, "to": 1000, "increment": .1}
            )
        
        self.resistencia.grid(row=2, column=0)
        
        c_frame = self._add_frame("Consola")
        self.consola = st.ScrolledText(c_frame, width = 75, height= 10)
        self.consola.configure(state='disabled',background="whitesmoke")
        self.consola.grid(sticky=tk.W, row=0, column=0)
        
        buttons = tk.Frame(self)
        buttons.grid(sticky=tk.W + tk.E, row=99)
        
        self.graphbutton = ttk.Button(
            buttons, text="Dibujar gráfica", command=self.master._on_plot)
        self.graphbutton.pack(side=tk.RIGHT)
        
        self.savebutton = ttk.Button(
            buttons, text="Guardar", command=self.master._on_save)
        self.savebutton.pack(side=tk.RIGHT)
        
        # self.stopbutton = ttk.Button(
        #     buttons, text="Parar", command=self.master.stopthread)
        # self.stopbutton.pack(side=tk.RIGHT)
        
        self.measurebutton = ttk.Button(
            buttons, text="Medir", command=self.master.threading)
        self.measurebutton.pack(side=tk.RIGHT)
        
        
        self._vars["Tipo de medida"].trace_add('write',self._show_widgets)
        
        self._vars["Tipo de medida"].set("I-V Diodo")
        

    def _show_widgets(self, *_):
        if self._vars["Tipo de medida"].get()=="I-V Diodo":
            self._vars["VDD Min"].set(-2)
            self._vars["VDD Max"].set(2)
            self._vars["Incremento"].set(0.02)
            self._vars["Valor de R (Ohm)"].set(100)
            self.vddmax.label.config(text='VDD Máximo')
            self.vddmin.grid(row=0, column=0)
            self.vddmax.grid(row=0, column=1)
            self.incremento.grid(row=0, column=2)
            self.resistencia.grid(row=2, column=0)
            self.vgsmin.grid_forget()
            self.vgsmax.grid_forget()
            self.incrementovgs.grid_forget()
            self.consola.configure(state='normal')
            self.consola.delete('1.0', tk.END)
            self.consola.configure(state='disabled')
            self.master._console_print(self.consola,"Seleccionada medida I-V del diodo\n","green")
            self.master._console_print(self.consola,"Asegúrese de que el primer interruptor está en MOSFET/Diode\n","green")
            self.master._console_print(self.consola,"Asegúrese de que el segundo interruptor está en Diode\n","green")
            self.master.medida_output=""

        elif self._vars["Tipo de medida"].get()=="Id-Vds MOS":
            self._vars["VDD Min"].set(0)
            self._vars["VDD Max"].set(10)
            self._vars["Incremento"].set(0.2)
            self._vars["Valor de R (Ohm)"].set(100)
            self.vddmin.grid(row=0, column=0)
            self.vddmax.grid(row=0, column=1)
            self.incremento.grid(row=0, column=2)
            self.vgsmin.grid(row=1, column=0)
            self.vgsmax.grid(row=1, column=1)
            self.vddmin.label.config(text='VDS Mínimo')
            self.vddmax.label.config(text='VDS Máximo')
            self.vgsmin.label.config(text='VGS Mínimo')
            self.vgsmax.label.config(text='VGS Máximo')
            self.incrementovgs.grid(row=1, column=2)
            self._vars["VGS Min"].set(0)
            self._vars["VGS Max"].set(5)
            self._vars["IncrementoVGS"].set(0.5)
            self.resistencia.grid(row=2, column=0)
            self.consola.configure(state='normal')
            self.consola.delete('1.0', tk.END)
            self.consola.configure(state='disabled')
            self.master._console_print(self.consola,"Seleccionada medida Id-Vds del MOSFET\n","green")
            self.master._console_print(self.consola,"Asegúrese de que el primer interruptor está en MOSFET/Diode\n","green")
            self.master._console_print(self.consola,"Asegúrese de que el segundo interruptor está en MOSFET\n","green")
            self.master.medida_output=""
            
        elif self._vars["Tipo de medida"].get()=="Id-Vgs MOS":
            self._vars["VDD Min"].set(1)
            self._vars["VDD Max"].set(5)
            self._vars["Incremento"].set(1)
            self._vars["VGS Min"].set(-2)
            self._vars["VGS Max"].set(5)
            self._vars["IncrementoVGS"].set(0.05)
            self._vars["Valor de R (Ohm)"].set(100)
            self.vddmin.label.config(text='VDS Mínimo')
            self.vddmax.label.config(text='VDS Máximo')
            self.vgsmin.label.config(text='VGS Mínimo')
            self.vgsmax.label.config(text='VGS Máximo')
            self.vddmin.grid(row=0, column=0)
            self.vddmax.grid(row=0, column=1)
            self.incremento.grid(row=0, column=2)      
            self.vgsmin.grid(row=1, column=0)
            self.vgsmax.grid(row=1, column=1)
            self.incrementovgs.grid(row=1, column=2)
            self.resistencia.grid(row=2, column=0)
            self.consola.configure(state='normal')
            self.consola.delete('1.0', tk.END)
            self.consola.configure(state='disabled')
            self.master._console_print(self.consola,"Seleccionada medida Id-Vgs del MOSFET\n","green")
            self.master._console_print(self.consola,"Asegúrese de que el primer interruptor está en MOSFET/Diode\n","green")
            self.master._console_print(self.consola,"Asegúrese de que el segundo interruptor está en MOSFET\n","green")
            self.master.medida_output=""

        elif self._vars["Tipo de medida"].get()=="Ic-Vce BJT":
            self._vars["VDD Min"].set(0)
            self._vars["VDD Max"].set(5)
            self._vars["Incremento"].set(0.1)
            self._vars["Valor de R (Ohm)"].set(100)
            self.vddmin.grid(row=0, column=0)
            self.vddmax.grid(row=0, column=1)
            self.incremento.grid(row=0, column=2)
            self.vgsmin.grid(row=1, column=0)
            self.vgsmax.grid(row=1, column=1)
            self.vddmin.label.config(text='VCE Mínimo')
            self.vddmax.label.config(text='VCE Máximo')
            self.vgsmin.label.config(text='IB (µA) Mínima')
            self.vgsmax.label.config(text='IB (µA) Máxima')
            self.incrementovgs.grid(row=1, column=2)
            self._vars["VGS Min"].set(0)
            self._vars["VGS Max"].set(50)
            self._vars["IncrementoVGS"].set(10)
            self.resistencia.grid_forget()
            self.consola.configure(state='normal')
            self.consola.delete('1.0', tk.END)
            self.consola.configure(state='disabled')
            self.master._console_print(self.consola,"Seleccionada medida Ic-Vce del BJT\n","green")
            self.master._console_print(self.consola,"Asegúrese de que el primer interruptor está en BJT\n","green")
            self.master.medida_output=""

        
class Application(tk.Tk):
    """Aplicación raíz"""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.title("USAL myDAQ - Medida de dispositivos")
        self.columnconfigure(0, weight=1)
        ttk.Label(
            self, text="USAL myDAQ - Medida de dispositivos",
            font=("TkDefaultFont", 16)
            ).grid(row=0)
        self.columnconfigure(0, weight=1)
        ttk.Label(
            self, text="Universidad de Salamanca, Licencia CC BY-NC-SA",
            font=("TkDefaultFont", 12)
            ).grid(row=1)
        self.recordform = DataRecordForm(self)
        self.recordform.grid(row=2, padx=10, sticky=(tk.W + tk.E))
        self.status = tk.StringVar()
        ttk.Label(
            self, textvariable=self.status
            ).grid(sticky=(tk.W + tk.E), row=3, padx=10)

        self._records_saved = 0
        
        self.system = nidaqmx.system.System.local()
        self._checkmyDAQ()
        
        self.medida_output=[]
        
    def _console_print(self,box,text,*color):
        box.configure(state='normal')
        if color:
            tagname=str(color)
            box.tag_config(tagname, foreground=color)
            box.insert("insert",text,tagname)
        else:
            box.insert("insert",text)

        box.see(tk.END)
        box.configure(state='disabled')
        
    def _checkmyDAQ(self):
        if len(self.system.devices)>0:
            self.status.set("Dispositivo encontrado: {}".format(self.system.devices[0].name))
            self._is_device = True
        else:
            self.status.set("No se encontraron dispositivos")
            self._is_device = False
        
    def _on_plot(self):
        
        class NavigationToolbar(NavigationToolbar2Tk):
       
            toolitems = [t for t in NavigationToolbar2Tk.toolitems if t[0] in ('Home', 'Forward', 'Back', 'Pan', 'Zoom', 'Save')]

        plt.rcParams["figure.dpi"] = 120
 
        if self.medida_output:

            
            def _replot():

                ax.clear()
                ax.set_axisbelow(True)
                ax.grid(visible=True, which='major', color='gainsboro', linestyle='-')
                canvas.get_tk_widget().pack_forget()                
                _plot()
                
            def _plot():
                                
                xdata=[]
                ydata=[]
            
                if tipomedida=="I-V Diodo":
                    xdata=[d["Vpn (V)"] for d in medidaploteada]
                    ydata=[d["Id (mA)"] for d in medidaploteada]        
                    
                    xdata=np.delete(np.array(xdata),0).astype(np.float32)
                    ydata=np.delete(np.array(ydata),0).astype(np.float32)       
                    
                    if tipografica.get() == "Línea":
                        ax.plot(xdata,ydata)
                    else:
                        ax.scatter(xdata,ydata,s=20)
                        
                    ax.set_xlabel("$V_{pn}$ (V)")
                    ax.set_ylabel("$I_d$ (mA)")
                    fig.subplots_adjust(right=0.85)
                    
                elif tipomedida=="Id-Vds MOS":
                    for i in range(len(medidaploteada)):
                        xdata=[d["VDS (V)"] for d in medidaploteada[i]]
                        ydata=[d["ID (mA)"] for d in medidaploteada[i]]        
                        
                        xdata=np.delete(np.array(xdata),0).astype(np.float32)
                        ydata=np.delete(np.array(ydata),0).astype(np.float32)       

                        try:
                            vgsvalue = "%.2f" % float(medidaploteada[i][1]["VGS (V)"])
                        except:
                            pass
                        
                        if tipografica.get() == "Línea":
                            ax.plot(xdata,ydata,label=f"VGS = {vgsvalue} V")
                        else:
                            ax.scatter(xdata,ydata,s=20,label=f"VGS = {vgsvalue} V")
                            
                    ax.set_xlabel("$V_{DS}$ (V)")
                    ax.set_ylabel("$I_D$ (mA)")
                    ax.legend(loc='center left', bbox_to_anchor=(1, 0.5))
                    fig.subplots_adjust(right=0.73)
                    
                elif tipomedida=="Id-Vgs MOS":
                    for i in range(len(medidaploteada)):
                        xdata=[d["VGS (V)"] for d in medidaploteada[i]]
                        ydata=[d["ID (mA)"] for d in medidaploteada[i]]        
                        
                        xdata=np.delete(np.array(xdata),0).astype(np.float32)
                        ydata=np.delete(np.array(ydata),0).astype(np.float32)       

                        try:
                            vdsvalue = "%.2f" % float(medidaploteada[i][1]["VDS (V)"])
                        except:
                            pass
                        
                        if tipografica.get() == "Línea":
                            ax.plot(xdata,ydata,label=f"VDS = {vdsvalue} V")
                        else:
                            ax.scatter(xdata,ydata,s=20,label=f"VDS = {vdsvalue} V")
                            
                    ax.set_xlabel("$V_{GS}$ (V)")
                    ax.set_ylabel("$I_D$ (mA)")
                    ax.legend(loc='center left', bbox_to_anchor=(1, 0.5))
                    fig.subplots_adjust(right=0.73)

                elif tipomedida=="Ic-Vce BJT":
                    for i in range(len(medidaploteada)):
                        xdata=[d["VCE (V)"] for d in medidaploteada[i]]
                        ydata=[d["IC (mA)"] for d in medidaploteada[i]]        
                        
                        xdata=np.delete(np.array(xdata),0).astype(np.float32)
                        ydata=np.delete(np.array(ydata),0).astype(np.float32)       
                        try:
                            ibvalue = "%.2f" % float(medidaploteada[i][1]["IB (µA)"])
                        except:
                            pass
                        if tipografica.get() == "Línea":
                            ax.plot(xdata,ydata,label=f"IB = {ibvalue} µA")
                        else:
                            ax.scatter(xdata,ydata,s=20,label=f"IB = {ibvalue} µA")
                            
                    ax.set_xlabel("$V_{CE}$ (V)")
                    ax.set_ylabel("$I_C$ (mA)")
                    ax.legend(loc='center left', bbox_to_anchor=(1, 0.5))
                    fig.subplots_adjust(right=0.73)
                    
                canvas.get_tk_widget().grid(sticky=tk.W + tk.E, row=2)
                canvas.draw()
            
            popup = tk.Toplevel(app)          
                        
            medidaploteada = self.medida_output
            tipomedida = self.recordform._vars["Tipo de medida"].get()
            tipodispo = self.recordform._vars["Ref"].get()
            
            tipografica = tk.StringVar()
            frame_window = tk.Frame(popup)
            frame_window.grid(sticky=tk.W, row=0)
            buttons = tk.Frame(frame_window)
            buttons.grid(sticky=tk.W,padx=10,row=1, column=0,columnspan=1)

            savebutton = ttk.Button(
                buttons, text="Guardar datos", command= lambda : self._on_savedata(medidaploteada,tipomedida,tipodispo))
            savebutton.grid(row=0)
            
            LabelInput(
                frame_window, "Tipo de gráfica", input_class=ttk.Radiobutton,
                var=tipografica,
                input_args={"values": ["Línea", "Puntos"]}
                ).grid(sticky=tk.W, padx=10, row=0, column=0,columnspan=1)
             
            fig = plt.Figure()
            canvas = FigureCanvasTkAgg(fig, master=frame_window)
            frame_toolbar = tk.Frame(popup)
            toolbar = NavigationToolbar(canvas,frame_toolbar)
            toolbar.update()
            frame_toolbar.grid(sticky=tk.W + tk.E, row=99)
            tipografica.set("Línea")

            ax = fig.add_subplot(111)
            ax.set_axisbelow(True)
            ax.grid(visible=True, which='major', color='gainsboro', linestyle='-')

            _plot()
            
            tipografica.trace_add('write',lambda *args: _replot())
   
            #tk.Button(popup, text="Cerrar la ventana", command=popup.destroy).grid(row=2)

        else:
            self._console_print(self.recordform.consola,"No hay medidas para representar\n","red")

    def _on_savedata(self,data,reftipo,refdispo):
        """Guardar archivo"""
        
        if data:
            datestring = datetime.today().strftime("%Y-%m-%d")
            
            files = [('Archivo separado por comas', '*.csv'),('Archivo de texto', '*.txt'),('Todos los archivos', '*.*')]
            if refdispo != "":      
                prename = "{}-{}-{}".format(refdispo,datestring,reftipo)
            else:
                prename = "{}-{}".format(datestring,reftipo)
                
            filename = asksaveasfilename(filetypes = files, defaultextension = files, initialfile = prename)

            if reftipo == "Id-Vds MOS":
                medidalist=[]
                mkeys=[]
                for k in range(len(data)):
                    mkeys.append(f"VGS (V) {k+1}")
                    mkeys.append(f"VDS (V) {k+1}")
                    mkeys.append(f"ID (mA) {k+1}")
                    
                for i in range(len(data[0])):  

                    medidasingle = []                      
                    for j in range(len(data)):
                        medidasingle.extend(list(data[j][i].values()))
                    medidalist.append(dict(zip(mkeys,medidasingle)))

            elif reftipo == "Ic-Vgs MOS":
                medidalist=[]
                mkeys=[]
                for k in range(len(data)):
                    mkeys.append(f"VDS (V) {k+1}")
                    mkeys.append(f"VGS (V) {k+1}")
                    mkeys.append(f"ID (mA) {k+1}")
                    
                for i in range(len(data[0])):  

                    medidasingle = []                      
                    for j in range(len(data)):
                        medidasingle.extend(list(data[j][i].values()))
                    medidalist.append(dict(zip(mkeys,medidasingle)))
                    
            elif reftipo == "Ic-Vce BJT":
                medidalist=[]
                mkeys=[]
                for k in range(len(data)):
                    mkeys.append(f"IB (µA) {k+1}")
                    mkeys.append(f"VCE (V) {k+1}")
                    mkeys.append(f"IC (mA) {k+1}")
                    
                for i in range(len(data[0])):  

                    medidasingle = []                      
                    for j in range(len(data)):
                        medidasingle.extend(list(data[j][i].values()))
                    medidalist.append(dict(zip(mkeys,medidasingle)))
            else:
                medidalist = data
                
            
            if filename!="":
                try:
                    a_file = open(filename, 'w', newline='')
                    if refdispo!="":
                        a_file.write(f"Dispositivo: {refdispo}\n")
                    else:
                        a_file.write("Dispositivo sin referencia\n")
                        
                    keys = medidalist[0].keys()
                    dict_writer = csv.DictWriter(a_file, keys, delimiter=";")
                    #dict_writer.writeheader()
                    dict_writer.writerows(medidalist)
                    a_file.close()
                    self._console_print(self.recordform.consola,"Archivo guardado con éxito\n","green")
    
                except:
                    self._console_print(self.recordform.consola,"Error al guardar el archivo\n","red")
                    self._console_print(self.recordform.consola,"Compruebe que no está abierto por otra aplicación\n","red")
            else:
                self._console_print(self.recordform.consola,"No se ha guardado la medida\n","red")
        else:
            self._console_print(self.recordform.consola,"No hay datos que guardar\n","red")


        
    def _on_save(self):
        """Guardar archivo"""
        
        if self.medida_output:
            datestring = datetime.today().strftime("%Y-%m-%d")
            reftipo = self.recordform._vars["Tipo de medida"].get()
            refdispo = self.recordform._vars["Ref"].get()
            files = [('Archivo separado por comas', '*.csv'),('Archivo de texto', '*.txt'),('Todos los archivos', '*.*')]
            if refdispo != "":      
                prename = "{}-{}-{}".format(refdispo,datestring,reftipo)
            else:
                prename = "{}-{}".format(datestring,reftipo)
                
            filename = asksaveasfilename(filetypes = files, defaultextension = files, initialfile = prename)

            if self.recordform._vars["Tipo de medida"].get() == "Id-Vds MOS":
                medidalist=[]
                mkeys=[]
                for k in range(len(self.medida_output)):
                    mkeys.append(f"VGS (V) {k+1}")
                    mkeys.append(f"VDS (V) {k+1}")
                    mkeys.append(f"ID (mA) {k+1}")
                    
                for i in range(len(self.medida_output[0])):  

                    medidasingle = []                      
                    for j in range(len(self.medida_output)):
                        medidasingle.extend(list(self.medida_output[j][i].values()))
                    medidalist.append(dict(zip(mkeys,medidasingle)))

            elif self.recordform._vars["Tipo de medida"].get() == "Id-Vgs MOS":
                medidalist=[]
                mkeys=[]
                for k in range(len(self.medida_output)):
                    mkeys.append(f"VDS (V) {k+1}")
                    mkeys.append(f"VGS (V) {k+1}")
                    mkeys.append(f"ID (mA) {k+1}")
                    
                for i in range(len(self.medida_output[0])):  

                    medidasingle = []                      
                    for j in range(len(self.medida_output)):
                        medidasingle.extend(list(self.medida_output[j][i].values()))
                    medidalist.append(dict(zip(mkeys,medidasingle)))

            elif self.recordform._vars["Tipo de medida"].get() == "Ic-Vce BJT":
                medidalist=[]
                mkeys=[]
                for k in range(len(self.medida_output)):
                    mkeys.append(f"IB (µA) {k+1}")
                    mkeys.append(f"VCE (V) {k+1}")
                    mkeys.append(f"IC (mA) {k+1}")
                    
                for i in range(len(self.medida_output[0])):  

                    medidasingle = []                      
                    for j in range(len(self.medida_output)):
                        medidasingle.extend(list(self.medida_output[j][i].values()))
                    medidalist.append(dict(zip(mkeys,medidasingle)))
            else:
                medidalist = self.medida_output
                
            
            if filename!="":
                try:
                    a_file = open(filename, 'w', newline='')
                    if refdispo!="":
                        a_file.write(f"Dispositivo: {refdispo}\n")
                    else:
                        a_file.write("Dispositivo sin referencia\n")
                        
                    keys = medidalist[0].keys()
                    dict_writer = csv.DictWriter(a_file, keys, delimiter=";")
                    #dict_writer.writeheader()
                    dict_writer.writerows(medidalist)
                    a_file.close()
                    self._console_print(self.recordform.consola,"Archivo guardado con éxito\n","green")
    
                except:
                    self._console_print(self.recordform.consola,"Error al guardar el archivo\n","red")
                    self._console_print(self.recordform.consola,"Compruebe que no está abierto por otra aplicación\n","red")
            else:
                self._console_print(self.recordform.consola,"No se ha guardado la medida\n","red")
        else:
            self._console_print(self.recordform.consola,"No hay datos que guardar\n","red")

        
    def threading(self):
        # Call work function
        self.t1=threading.Thread(target=self._on_run)
        self.t1.setDaemon(True)
        self.t1.start()        

    # def stopthread(self):
    #     if self.t1:
    #         self.t1._stop()
            
    def _on_run(self):
        """Ejecución de las medidas"""

        if self._is_device:
            
            if self.recordform._vars["Tipo de medida"].get()=="I-V Diodo":
                self._IVdiode_measure()
            elif self.recordform._vars["Tipo de medida"].get()=="Id-Vds MOS":
                self._IVMOS_measure()
            elif self.recordform._vars["Tipo de medida"].get()=="Id-Vgs MOS":
                self._IVGMOS_measure()
            elif self.recordform._vars["Tipo de medida"].get()=="Ic-Vce BJT":
                self._IVBJT_measure()
    
    def _writemyDAQ(self,channel,value,*_):
        with nidaqmx.Task() as task:
            task.ao_channels.add_ao_voltage_chan('{}/{}'.format(self.system.devices[0].name,channel))
            task.write(value)
            task.wait_until_done()
            
    def _readmyDAQ(self,channel,*_):
        with nidaqmx.Task() as task:
            task.ai_channels.add_ai_voltage_chan('{}/{}'.format(self.system.devices[0].name,channel))
            return task.read()
        
    def _IVdiode_measure(self):
        rangemin = self.recordform._vars['VDD Min'].get()
        rangemax = self.recordform._vars['VDD Max'].get()
        incr = self.recordform._vars['Incremento'].get()
        resistor = self.recordform._vars['Valor de R (Ohm)'].get()

        if incr != 0:
            try:
                numvdd = int((rangemax-rangemin)/incr)+1
                vdd = np.linspace(rangemin, rangemax, numvdd)
            except:
                self._console_print(self.recordform.consola,"Revise los parámetros elegidos\n",'red')
                return
        else:
            if rangemax != rangemin:
                self._console_print(self.recordform.consola,"Revise los parámetros elegidos\n",'red')
                return
            else:
                vdd = [rangemax]
                
        self._console_print(self.recordform.consola,"Iniciando medida\n",'blue')      
        self.medida_output = [{"VDD (V)": "VDD (V)", "Vpn (V)": "Vpn (V)", "Id (mA)": "Id (mA)"}]
        
        countervdd = 0
        
        while countervdd < len(vdd):
            value = vdd[countervdd]
            self._writemyDAQ("ao0", value)
            vpn = self._readmyDAQ("ai0")
           
            valim = "%.4f" % round(vpn,4)
            vdiodo = "%.4f" % round(value,4)
            current = "%.4f" % round(((vpn-value)/resistor*1000),4)
            if abs(float(current))*30 <= 500:   #Se comprueba en relación a la potencia total disponible (500 mw) en los +-15 (30)     
                lectura = 'VDD (V): '+str(valim)+ \
                    ' ; Vpn (V): '+str(vdiodo)+ \
                        ' ; ID (mA): '+ str(current)+'\n'
                self._console_print(self.recordform.consola,lectura)
                self.medida_output.append({"VDD (V)": valim, "Vpn (V)": vdiodo, "Id (mA)": current})
                countervdd = countervdd + 1
            else:
                self._console_print(self.recordform.consola,"Excedida potencia máxima\n",'blue')
                countervdd = len(vdd)
        self._console_print(self.recordform.consola,"Medida finalizada\n",'blue')
        
    def _IVMOS_measure(self):
        
        rangemin = self.recordform._vars['VDD Min'].get()
        rangemax = self.recordform._vars['VDD Max'].get()
        incr = self.recordform._vars['Incremento'].get()
        resistor = self.recordform._vars['Valor de R (Ohm)'].get()

        rangevgsmin = self.recordform._vars['VGS Min'].get()
        rangevgsmax = self.recordform._vars['VGS Max'].get()
        incrvgs = self.recordform._vars['IncrementoVGS'].get()
        
        if incr != 0:
            try:
                numvdd = int((rangemax-rangemin)/incr)+1
                vdd = np.linspace(rangemin, rangemax, numvdd)
            except:
                self._console_print(self.recordform.consola,"Revise los parámetros elegidos\n",'red')
                return
        else:
            if rangemax != rangemin:
                self._console_print(self.recordform.consola,"Revise los parámetros elegidos\n",'red')
                return
            else:
                vdd = [rangemax]
                
        if incrvgs != 0:
            try:
                numvgs = int((rangevgsmax-rangevgsmin)/incrvgs)+1
                vgs = np.linspace(rangevgsmin, rangevgsmax, numvgs)
            except:
                self._console_print(self.recordform.consola,"Revise los parámetros elegidos\n",'red')
                return
        else:
            if rangevgsmax != rangevgsmin:
                self._console_print(self.recordform.consola,"Revise los parámetros elegidos\n",'red')
                return
            else:
                vgs = [rangevgsmax]

        self._console_print(self.recordform.consola,"Iniciando medida\n",'blue')
        self.medida_output =[]        
        countervgs = 0
        while countervgs < len(vgs):
            countervdd = 0
            valuevgs = vgs[countervgs]
            self._writemyDAQ("ao1", valuevgs)
            medida_given_vgs = [{"VGS (V)": "VGS (V)", "VDS (V)": "VDS (V)", "ID (mA)": "ID (mA)"}]
        
            while countervdd < len(vdd):
                value = vdd[countervdd]
                self._writemyDAQ("ao0", value)
                vpuerta = "%.4f" % round(valuevgs,4)
                vmeas = self._readmyDAQ("ai0")
                if vmeas < 10.5 and vmeas > -10.5:  # Se fija límite 10.5, máximo que se puede leer por AI
                    ids = (vmeas-value)/resistor*1000
                    vds = "%.4f" % round(value,4)
                    ids = "%.4f" % round(ids,4)
                    if abs(float(ids))*30 <= 500:   #Se comprueba en relación a la potencia total disponible (500 mw) en los +-15 (30)                      
                        lectura = 'VGS (V): '+str(vpuerta)+ \
                            ' ; VDS (V): '+str(vds)+ \
                                ' ; ID (mA): '+ str(ids)+'\n' 
                        self._console_print(self.recordform.consola,lectura)
                        medida_given_vgs.append({"VGS (V)": vpuerta, "VDS (V)": vds, "ID (mA)": ids})
                        countervdd = countervdd + 1
                    else:
                        self._console_print(self.recordform.consola,"Excedida potencia máxima\n",'blue')
                        countervdd = len(vdd)
                else:
                    countervdd = countervdd + 1
            self.medida_output.append(medida_given_vgs)
            countervgs = countervgs + 1    

        self._console_print(self.recordform.consola,"Medida finalizada\n",'blue')

    def _IVGMOS_measure(self):
        rangemin = self.recordform._vars['VDD Min'].get()
        rangemax = self.recordform._vars['VDD Max'].get()
        incr = self.recordform._vars['Incremento'].get()
        resistor = self.recordform._vars['Valor de R (Ohm)'].get()

        rangevgsmin = self.recordform._vars['VGS Min'].get()
        rangevgsmax = self.recordform._vars['VGS Max'].get()
        incrvgs = self.recordform._vars['IncrementoVGS'].get()
        
        if incr != 0:
            try:
                numvdd = int((rangemax-rangemin)/incr)+1
                vdd = np.linspace(rangemin, rangemax, numvdd)
            except:
                self._console_print(self.recordform.consola,"Revise los parámetros elegidos\n",'red')
                return
        else:
            if rangemax != rangemin:
                self._console_print(self.recordform.consola,"Revise los parámetros elegidos\n",'red')
                return
            else:
                vdd = [rangemax]
                
        if incrvgs != 0:
            try:
                numvgs = int((rangevgsmax-rangevgsmin)/incrvgs)+1
                vgs = np.linspace(rangevgsmin, rangevgsmax, numvgs)
            except:
                self._console_print(self.recordform.consola,"Revise los parámetros elegidos\n",'red')
                return
        else:
            if rangevgsmax != rangevgsmin:
                self._console_print(self.recordform.consola,"Revise los parámetros elegidos\n",'red')
                return
            else:
                vgs = [rangevgsmax]

        self._console_print(self.recordform.consola,"Iniciando medida\n",'blue')
        self.medida_output =[]        
        countervdd = 0

        while countervdd < len(vdd):
            countervgs = 0
            valuevds = vdd[countervdd]
            self._writemyDAQ("ao0", valuevds) 
            medida_given_vds = [{"VDS (V)": "VDS (V)", "VGS (V)": "VGS (V)", "ID (mA)": "ID (mA)"}]
        
            while countervgs < len(vgs):
                valuevgs = vgs[countervgs]
                self._writemyDAQ("ao1", valuevgs)
                vpuerta = "%.4f" % round(valuevgs,4)
                vmeas = self._readmyDAQ("ai0")
                if vmeas < 10.5 and vmeas > -10.5:  # Se fija límite 10.5, máximo que se puede leer por AI
                    ids = (vmeas-valuevds)/resistor*1000
                    vds = "%.4f" % round(valuevds,4)
                    ids = "%.4f" % round(ids,4)
                    if abs(float(ids))*30 <= 500:   #Se comprueba en relación a la potencia total disponible (500 mw) en los +-15 (30)                      
                        lectura = 'VDS (V): '+str(vds)+ \
                            ' ; VGS (V): '+str(vpuerta)+ \
                                ' ; ID (mA): '+ str(ids)+'\n' 
                        self._console_print(self.recordform.consola,lectura)
                        medida_given_vds.append({"VDS (V)": vds, "VGS (V)": vpuerta, "ID (mA)": ids})
                        countervgs = countervgs + 1
                    else:
                        self._console_print(self.recordform.consola,"Excedida potencia máxima\n",'blue')
                        countervgs = len(vgs)
                else:
                    countervgs = countervgs + 1
            self.medida_output.append(medida_given_vds)
            countervdd = countervdd + 1  
            
        self._console_print(self.recordform.consola,"Medida finalizada\n",'blue')
        
    def _IVBJT_measure(self):
        
        rangemin = self.recordform._vars['VDD Min'].get()
        rangemax = self.recordform._vars['VDD Max'].get()
        incr = self.recordform._vars['Incremento'].get()
        resistor = 10

        rangevgsmin = self.recordform._vars['VGS Min'].get()
        rangevgsmax = self.recordform._vars['VGS Max'].get()
        incrvgs = self.recordform._vars['IncrementoVGS'].get()

        if incr != 0:
            try:
                numvdd = int((rangemax-rangemin)/incr)+1
                vdd = np.linspace(rangemin, rangemax, numvdd)
            except:
                self._console_print(self.recordform.consola,"Revise los parámetros elegidos\n",'red')
                return
        else:
            if rangemax != rangemin:
                self._console_print(self.recordform.consola,"Revise los parámetros elegidos\n",'red')
                return
            else:
                vdd = [rangemax]
                
        if incrvgs != 0:
            try:
                numvgs = int((rangevgsmax-rangevgsmin)/incrvgs)+1
                vgs = np.linspace(rangevgsmin, rangevgsmax, numvgs)*0.1
            except:
                self._console_print(self.recordform.consola,"Revise los parámetros elegidos\n",'red')
                return
        else:
            if rangevgsmax != rangevgsmin:
                self._console_print(self.recordform.consola,"Revise los parámetros elegidos\n",'red')
                return
            else:
                vgs = [rangevgsmax]*0.1

        self._console_print(self.recordform.consola,"Iniciando medida\n",'blue')
        self.medida_output =[]        
        countervgs = 0
        while countervgs < len(vgs):
            countervdd = 0
            valuevgs = vgs[countervgs]
            self._writemyDAQ("ao0", valuevgs)
            medida_given_vgs = [{"IB (µA)": "IB (µA)", "VCE (V)": "VCE (V)", "IC (mA)": "IC (mA)"}]
        
            while countervdd < len(vdd):
                value = vdd[countervdd]
                self._writemyDAQ("ao1", value)
                vpuerta = "%.4f" % round(valuevgs*10,4)
                vmeas = self._readmyDAQ("ai0")
                vemitter = self._readmyDAQ("ai1")
                ids = (value-vmeas)/resistor*1000
                vds = "%.4f" % round(vmeas-vemitter,4)
                ids = "%.4f" % round(ids,4)
                if abs(float(ids))*30 <= 500:   #Se comprueba en relación a la potencia total disponible (500 mw) en los +-15 (30)      
                    lectura = 'IB (µA): '+str(vpuerta)+ \
                     ' ; VCE (V): '+str(vds)+ \
                     ' ; IC (mA): '+ str(ids)+'\n' 
                    self._console_print(self.recordform.consola,lectura)
                    medida_given_vgs.append({"IB (µA)": vpuerta, "VCE (V)": vds, "IC (mA)": ids})
                    countervdd = countervdd + 1
                else:
                    self._console_print(self.recordform.consola,"Excedida potencia máxima\n",'blue')
                    countervdd = len(vdd)  
            self.medida_output.append(medida_given_vgs)
            countervgs = countervgs + 1    

        self._console_print(self.recordform.consola,"Medida finalizada\n",'blue')
        
if __name__ == "__main__":
    app = Application()
    #app.iconbitmap('D:\\beta\\ICONO.ico')
    app.mainloop()