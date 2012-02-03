module Excelator

# excelator_yema_v1-1.rb   -------   Ver Summer 2011 v1-1
#
# -------- Modo de Uso (Copie lo que sigue en su programa)
# En la primera instrucción del programa que va a utilizar esta yema colocar "require ./excelator_yema.rb"  
# f=Excelator.new("libro")      #Crea el libro excel de salida "libro.xml"
# f.iniciar         #Inicia el libro con los encabezados de excel y los estilos a utilizar
#     Luego se definen las planillas del libro
# f.iniciar_planilla("planilla")  #Inicia una planilla en el libro (hay que hacerlo para cada planilla
# f.columna(ancho)      #Hay que hacer uno para cada columna con el ancho 9999 medido en caracteres, todas se definene una debajo de la otra 
#     Luego se definen las lineas:
# f.iniciar_linea
# f.celda("tipo","contenido") #Hay que enviar las celdas en el orden que aparecen en la linea.
#             Tipos: Encabezado, GranTitulo, Titulo, Alfa, Num, Fecha, Total
#             Si en el contenido aparece un signo igual lo considera una formula  
# f.terminar_linea
#      Para terminar
# f.terminar_planilla   #Termina la planilla en curso del Libro
# f.terminar
#
class Excelator < File

  def iniciar 
    self.puts "<?xml version=\"1.0\"?>"
    self.puts "<?mso-application progid=\"Excel.Sheet\"?>"
    self.puts "<ss:Workbook xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\">"
    self.puts "<ss:Styles>"
    self.puts "<ss:Style ss:ID=\"Encabezado\">"
    self.puts "<ss:Font ss:Bold=\"1\" ss:Size=\"16\"/>"
    self.puts "</ss:Style>"
    self.puts "<ss:Style ss:ID=\"Pie\">"
    self.puts "<ss:Font ss:Size=\"7\"/>"
    self.puts "</ss:Style>"
    self.puts "<ss:Style ss:ID=\"GranTitulo\">"
    self.puts "<ss:Font ss:Bold=\"1\" ss:Size=\"11\"/>"
    self.puts "<ss:Alignment ss:Horizontal=\"Center\"/>"
    self.puts "</ss:Style>"
    self.puts "<ss:Style ss:ID=\"Titulo\">"
    self.puts "<ss:Font ss:Bold=\"1\" ss:Italic=\"1\" ss:Size=\"10\"/>"
    self.puts "<ss:Alignment ss:Horizontal=\"Center\"/>"
    self.puts "<ss:Interior ss:Color=\"#e6e6e6\" ss:Pattern=\"Solid\"/>"
    self.puts "</ss:Style>"
    self.puts "<ss:Style ss:ID=\"Num\">"
    self.puts "<ss:Font ss:Size=\"10\"/>"
    self.puts "<ss:Alignment ss:Horizontal=\"Right\"/>"
    self.puts "</ss:Style>"
    self.puts "<ss:Style ss:ID=\"Alfa\">"
    self.puts "<ss:Font ss:Size=\"10\"/>"
    self.puts "<ss:Alignment ss:Horizontal=\"Left\"/>"
    self.puts "</ss:Style>"
    self.puts "<ss:Style ss:ID=\"Fecha\">"
    self.puts "<ss:Font ss:Size=\"10\"/>"
    self.puts "<ss:Alignment ss:Horizontal=\"Center\"/>"
    self.puts "<ss:NumberFormat ss:Format=\"Short Date\"/>"
    self.puts "</ss:Style>"
    self.puts "<ss:Style ss:ID=\"Total\">"
    self.puts "<ss:Font ss:Bold=\"1\" ss:Size=\"10\"/>"   
    self.puts "<ss:Alignment ss:Horizontal=\"Right\"/>"
    self.puts "<ss:NumberFormat ss:Format=\"0.00\"/>"
    self.puts "</ss:Style>"
    self.puts "<ss:Style ss:ID=\"Cifra\">"
    self.puts "<ss:Font ss:Size=\"10\"/>"
    self.puts "<ss:Alignment ss:Horizontal=\"Right\"/>"
    self.puts "<ss:NumberFormat ss:Format=\"0.00\"/>"
    self.puts "</ss:Style>"
    self.puts "</ss:Styles>"
  end
  
  def terminar
    self.puts "</ss:Workbook>"
  end

  def iniciar_planilla(planilla)
    self.puts "<ss:Worksheet ss:Name=\"#{planilla}\">"
    self.puts "<ss:Table>"
  end

  def terminar_planilla(autofiltro=false, filini, filfin, colfin, defaultcol, defaultvalor, frizado, lineafrizer)
    3.times{self.puts"<ss:Row></ss:Row>"}
    self.puts "<ss:Row>"
    self.puts "<ss:Cell ss:StyleID=\"Pie\">"
    self.puts "<ss:Data ss:Type=\"String\">------- Estudio Urien y Asociados - MAZARS - Quintana 585, CABA - Tel: 4804-6502 -------</ss:Data>"
    self.puts "</ss:Cell>"
    self.puts "</ss:Row>"
    self.puts "</ss:Table>"
    if autofiltro then
      self.puts "<x:AutoFilter x:Range=\"R#{filini.to_s}C1:R#{filfin.to_s}C#{colfin.to_s}\">"
      if (defaultcol!=-1) then
        self.puts "<x:AutoFilterColumn x:Index=\"1\" x:Type=\"Custom\">"
        self.puts "<x:AutoFilterCondition x:Operator=\"Equals\" x:Value=\"#{defaultvalor}\"/>"
        self.puts "</x:AutoFilterColumn>"
      end
      self.puts "</x:AutoFilter>"
    end
    if frizado then
      self.puts "<WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\">"
      self.puts "<FreezePanes/>"
      self.puts "<FrozenNoSplit/>"
      self.puts "<SplitHorizontal>#{lineafrizer}</SplitHorizontal>"
      self.puts "<TopRowBottomPane>#{lineafrizer}</TopRowBottomPane>"
      self.puts "<ActivePane>2</ActivePane>"
      self.puts "<Panes>"
      self.puts "<Pane>"
      self.puts "<Number>3</Number>"
      self.puts "</Pane>"
      self.puts "<Pane>"
      self.puts "<Number>2</Number>"
      self.puts "<RangeSelection>R6</RangeSelection>"
      self.puts "</Pane>"
      self.puts "</Panes>"
      self.puts "</WorksheetOptions>"
    end
    self.puts "</ss:Worksheet>"
  end

  def columna(ancho)
    self.puts "<ss:Column ss:Width=\"#{(ancho*8)+1}\"/>"
  end

  def iniciar_linea
    self.puts "<ss:Row>"
  end 

  def terminar_linea
    self.puts "</ss:Row>"
  end 

  def autofiltro(filfin,colfin)
    self.puts "<AutoFilter x:Range=\"R6C1:R#{filfin.to_s}C#{colfin.to_s}\" xmlns=\"urn:schemas-microsoft-com:office:excel\">"
    self.puts "</AutoFilter>"
  end 

  def celda(tipo="Alfa",contenido="")
    self.write "<ss:Cell ss:StyleID=\"#{tipo}\""
    formula=/=/.match(contenido)
    self.write " ss:Formula=\"#{contenido}\"" if formula
    self.write ">"
    self.write "<ss:Data " 
    self.write "ss:Type=\"String\">#{contenido}" if (tipo!="Num" and tipo!="Cifra" and tipo!="Total" and tipo!="Fecha" and !formula)
    self.write "ss:Type=\"Number\">#{contenido}" if (tipo=="Num" or tipo=="Cifra")
    self.write "ss:Type=\"String\">#{contenido}" if (formula or tipo=="Total")
    self.write "ss:Type=\"DateTime\">#{contenido}" if (tipo=="Fecha")
    self.write "</ss:Data>"
    self.puts "</ss:Cell>"
  end
  
  def Frizado(linea)
  end

#Class end
end

#Module end
end