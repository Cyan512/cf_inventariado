import ExcelJS from 'exceljs'
import { saveAs } from 'file-saver'
import { useState, useEffect } from 'react'
import { supabase } from './utils/supabase'

interface Producto {
  id: number
  nombre: string
  cantidad: number
  image_url: string
}

interface ProductoTemporal {
  id: string
  nombre: string
  cantidad: number
  file: File
  preview: string
}

export default function App() {
  const [productos, setProductos] = useState<Producto[]>([])
  const [imagenesSeleccionadas, setImagenesSeleccionadas] = useState<ProductoTemporal[]>([])
  const [loading, setLoading] = useState(false)
  const [showForm, setShowForm] = useState(false)

  async function getProductos() {
    const { data, error } = await supabase
      .from('productos')
      .select()
      .order('id', { ascending: false })

    if (error) {
      console.log(error)
      return
    }

    setProductos(data || [])
  }

  useEffect(() => {
    getProductos()
  }, [])

  async function exportarExcel() {
    const workbook = new ExcelJS.Workbook()
    const worksheet = workbook.addWorksheet('Productos')

    worksheet.columns = [
      { header: 'ID', key: 'id', width: 10 },
      { header: 'Nombre', key: 'nombre', width: 30 },
      { header: 'Cantidad', key: 'cantidad', width: 15 },
      { header: 'Imagen', key: 'imagen', width: 25 },
    ]

    for (let i = 0; i < productos.length; i++) {
      const producto = productos[i]

      const row = worksheet.addRow({
        id: producto.id,
        nombre: producto.nombre,
        cantidad: producto.cantidad,
      })

      row.height = 80

      if (producto.image_url) {
        const response = await fetch(producto.image_url)
        const blob = await response.blob()

        const arrayBuffer = await blob.arrayBuffer()

        const imageId = workbook.addImage({
          buffer: arrayBuffer,
          extension: 'png',
        })

        worksheet.addImage(imageId, {
          tl: { col: 3, row: i + 1 },
          ext: { width: 80, height: 80 },
        })
      }
    }

    const buffer = await workbook.xlsx.writeBuffer()

    saveAs(
      new Blob([buffer]),
      'productos.xlsx'
    )
  }

  async function subirImagenCloudinary(file: File) {
    const formData = new FormData() 

    formData.append('file', file)
    formData.append(
      'upload_preset',
      import.meta.env.VITE_CLOUDINARY_UPLOAD_PRESET
    )

    const res = await fetch(
      `https://api.cloudinary.com/v1_1/${import.meta.env.VITE_CLOUDINARY_CLOUD_NAME}/image/upload`,
      {
        method: 'POST',
        body: formData,
      }
    )

    const data = await res.json()
    return data.secure_url
  }

  function handleSeleccionarImagenes(e: React.ChangeEvent<HTMLInputElement>) {
    const files = e.target.files ? Array.from(e.target.files) : []
    
    const nuevos: ProductoTemporal[] = files.map((file) => ({
      id: Date.now().toString() + Math.random().toString(36).substr(2, 9),
      nombre: file.name.replace(/\.[^/.]+$/, ''),
      cantidad: 1,
      file,
      preview: URL.createObjectURL(file),
    }))

    setImagenesSeleccionadas([...imagenesSeleccionadas, ...nuevos])
  }

  function actualizarProducto(id: string, campo: 'nombre' | 'cantidad', valor: string | number) {
    setImagenesSeleccionadas(
      imagenesSeleccionadas.map((p) =>
        p.id === id ? { ...p, [campo]: valor } : p
      )
    )
  }

  function eliminarProducto(id: string) {
    setImagenesSeleccionadas(imagenesSeleccionadas.filter((p) => p.id !== id))
  }

  async function agregarTodos() {
    if (imagenesSeleccionadas.length === 0) {
      alert('Selecciona al menos una imagen')
      return
    }

    setLoading(true)

    try {
      for (const prod of imagenesSeleccionadas) {
        const imageUrl = await subirImagenCloudinary(prod.file)

        const { error } = await supabase
          .from('productos')
          .insert([
            {
              nombre: prod.nombre,
              cantidad: prod.cantidad,
              image_url: imageUrl,
            },
          ])

        if (error) {
          console.log(error)
          alert('Error al agregar producto')
          return
        }
      }

      setImagenesSeleccionadas([])
      setShowForm(false)
      getProductos()
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="min-h-screen bg-gray-50">
      {/* Header */}
      <div className="bg-white border-b border-gray-200">
        <div className="px-6 py-4">
          <div className="flex items-center justify-between">
            <div>
              <h1 className="text-xl font-semibold text-gray-900">Sistema de Inventario</h1>
              <p className="text-sm text-gray-500 mt-0.5">Gestión de productos</p>
            </div>
            <button
              onClick={exportarExcel}
              className="px-4 py-2 bg-green-600 text-white text-sm font-medium hover:bg-green-700 transition-colors"
            >
              Exportar Excel
            </button>
            <button
              onClick={() => setShowForm(!showForm)}
              className="px-4 py-2 bg-blue-600 text-white text-sm font-medium hover:bg-blue-700 transition-colors"
            >
              {showForm ? 'Cancelar' : '+ Nuevo Producto'}
            </button>
          </div>
        </div>
      </div>

      <div className="px-6 py-6">
        {/* Formulario */}
        {showForm && (
          <div className="bg-white border border-gray-200 mb-6">
            <div className="px-6 py-4 border-b border-gray-200">
              <h2 className="text-base font-semibold text-gray-900">Agregar Producto</h2>
            </div>
            <div className="px-6 py-5">
              <div className="mb-5">
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Seleccionar Imágenes
                </label>
                <input
                  type="file"
                  accept="image/*"
                  multiple
                  onChange={handleSeleccionarImagenes}
                  className="hidden"
                  id="file-upload"
                />
                <label
                  htmlFor="file-upload"
                  className="flex items-center justify-center w-full px-3 py-3 border-2 border-dashed border-gray-300 bg-white text-sm text-gray-700 hover:bg-gray-50 cursor-pointer"
                >
                  <svg className="w-5 h-5 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16l4.586-4.586a2 2 0 012.828 0L16 16m-2-2l1.586-1.586a2 2 0 012.828 0L20 14m-6-6h.01M6 20h12a2 2 0 002-2V6a2 2 0 00-2-2H6a2 2 0 00-2 2v12a2 2 0 002 2z" />
                  </svg>
                  Seleccionar múltiples imágenes
                </label>
              </div>

              {imagenesSeleccionadas.length > 0 && (
                <div className="mt-6">
                  <div className="flex items-center justify-between mb-4">
                    <h3 className="text-sm font-semibold text-gray-900">
                      Productos ({imagenesSeleccionadas.length})
                    </h3>
                    <button
                      type="button"
                      onClick={agregarTodos}
                      disabled={loading}
                      className="px-4 py-2 bg-green-600 text-white text-sm font-medium hover:bg-green-700 disabled:opacity-50"
                    >
                      {loading ? 'Agregando...' : 'Agregar todos'}
                    </button>
                  </div>
                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                    {imagenesSeleccionadas.map((prod) => (
                      <div key={prod.id} className="border border-gray-200 p-3 relative">
                        <button
                          type="button"
                          onClick={() => eliminarProducto(prod.id)}
                          className="absolute top-2 right-2 text-red-500 hover:text-red-700"
                        >
                          <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
                          </svg>
                        </button>
                        <img
                          src={prod.preview}
                          alt={prod.nombre}
                          className="w-full h-32 object-cover mb-3 border border-gray-200"
                        />
                        <div className="space-y-2">
                          <div>
                            <label className="block text-xs text-gray-500 mb-1">Nombre</label>
                            <input
                              type="text"
                              value={prod.nombre}
                              onChange={(e) => actualizarProducto(prod.id, 'nombre', e.target.value)}
                              className="w-full px-2 py-1 border border-gray-300 text-sm"
                            />
                          </div>
                          <div>
                            <label className="block text-xs text-gray-500 mb-1">Cantidad</label>
                            <input
                              type="number"
                              value={prod.cantidad}
                              onChange={(e) => actualizarProducto(prod.id, 'cantidad', Number(e.target.value))}
                              className="w-full px-2 py-1 border border-gray-300 text-sm"
                            />
                          </div>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
          </div>
        )}

        {/* Tabla */}
        <div className="bg-white border border-gray-200">
          <div className="px-6 py-4 border-b border-gray-200">
            <div className="flex items-center justify-between">
              <h2 className="text-base font-semibold text-gray-900">Lista de Productos</h2>
              <span className="text-sm text-gray-500">{productos.length} registros</span>
            </div>
          </div>

          {productos.length === 0 ? (
            <div className="px-6 py-12 text-center">
              <p className="text-sm text-gray-500">No hay productos registrados</p>
            </div>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full">
                <thead>
                  <tr className="bg-gray-50 border-b border-gray-200">
                    <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                      ID
                    </th>
                    <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                      Imagen
                    </th>
                    <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                      Nombre del Producto
                    </th>
                    <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                      Cantidad
                    </th>
                  </tr>
                </thead>
                <tbody className="bg-white divide-y divide-gray-200">
                  {productos.map((producto) => (
                    <tr key={producto.id} className="hover:bg-gray-50 transition-colors">
                      <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">
                        #{producto.id}
                      </td>
                      <td className="px-6 py-4 whitespace-nowrap">
                        {producto.image_url ? (
                          <img
                            src={producto.image_url}
                            alt={producto.nombre}
                            className="h-12 w-12 object-cover border border-gray-200"
                          />
                        ) : (
                          <div className="h-12 w-12 bg-gray-100 border border-gray-200 flex items-center justify-center">
                            <span className="text-xs text-gray-400">N/A</span>
                          </div>
                        )}
                      </td>
                      <td className="px-6 py-4 whitespace-nowrap">
                        <div className="text-sm font-medium text-gray-900">{producto.nombre}</div>
                      </td>
                      <td className="px-6 py-4 whitespace-nowrap">
                        <div className="text-sm text-gray-900">{producto.cantidad}</div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </div>
    </div>
  )
}
