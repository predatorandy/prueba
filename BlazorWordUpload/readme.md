# Blazor WASM Word Document Uploader

Componente Blazor WebAssembly para subir documentos Word (.docx) y mostrar su contenido en formato tabla.

## Estructura del Proyecto

```
BlazorWordUpload/
├── Components/
│   └── WordDocumentUploader.razor    # Componente principal
├── Pages/
│   └── Home.razor                     # Página de inicio
├── wwwroot/
│   └── css/
│       └── app.css                    # Estilos
├── App.razor                          # Componente raíz
├── Routes.razor                       # Enrutamiento
├── MainLayout.razor                   # Layout principal
├── Program.cs                         # Punto de entrada
└── BlazorWordUpload.csproj            # Archivo de proyecto
```

## Características

- **Subida de archivos .docx**: Validación de tipo de archivo
- **Extracción de tablas**: Detecta automáticamente las tablas del documento
- **Extracción de párrafos**: Muestra el texto fuera de las tablas
- **Interfaz responsiva**: Diseño limpio con Bootstrap-like styles
- **Manejo de errores**: Mensajes claros para el usuario
- **Indicador de procesamiento**: Feedback visual durante la carga

## Uso del Componente

### En una página Blazor:

```razor
@page "/upload"
@using BlazorWordUpload.Components

<WordDocumentUploader />
```

### Cómo funciona:

1. El usuario selecciona un archivo .docx
2. El componente lee el archivo en el navegador (client-side)
3. Extrae las tablas y las muestra con encabezados
4. Extrae los párrafos de texto libre
5. Muestra todo formateado en una tabla HTML

## Ejecutar el Proyecto

```bash
cd BlazorWordUpload
dotnet restore
dotnet run
```

El aplicación estará disponible en `https://localhost:5001` o `http://localhost:5000`

## Dependencias

- .NET 10
- Microsoft.AspNetCore.Components.WebAssembly
- DocumentFormat.OpenXml 3.2.0

## Limitaciones

- Tamaño máximo de archivo: 10 MB
- Solo soporta formato .docx (no .doc antiguo)
- Procesa solo la primera tabla encontrada como estructura principal
