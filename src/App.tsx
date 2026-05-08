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

export default function App() {
  const [productos, setProductos] = useState<Producto[]>([])
  const [nombre, setNombre] = useState('')
  const [cantidad, setCantidad] = useState(0)
  const [imageFile, setImageFile] = useState<File | null>(null)
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

  async function agregarProducto(e: React.FormEvent<HTMLFormElement>) {
    e.preventDefault()

    if (!nombre.trim()) {
      alert('Ingresa el nombre del producto')
      return
    }

    if (!imageFile) {
      alert('Selecciona una imagen')
      return
    }

    setLoading(true)

    try {
      const imageUrl = await subirImagenCloudinary(imageFile)

      const { error } = await supabase
        .from('productos')
        .insert([
          {
            nombre,
            cantidad,
            image_url: imageUrl,
          },
        ])

      if (error) {
        console.log(error)
        alert('Error al agregar producto')
        return
      }

      setNombre('')
      setCantidad(0)
      setImageFile(null)
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
            <form onSubmit={agregarProducto} className="px-6 py-5">
              <div className="grid grid-cols-3 gap-5 mb-5">
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    Nombre del Producto
                  </label>
                  <input
                    type="text"
                    value={nombre}
                    onChange={(e) => setNombre(e.target.value)}
                    className="w-full px-3 py-2 border border-gray-300 text-sm focus:outline-none focus:border-blue-500"
                    placeholder="Ingrese el nombre"
                  />
                </div>

                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    Cantidad
                  </label>
                  <input
                    type="number"
                    value={cantidad}
                    onChange={(e) => setCantidad(Number(e.target.value))}
                    className="w-full px-3 py-2 border border-gray-300 text-sm focus:outline-none focus:border-blue-500"
                    placeholder="0"
                  />
                </div>

                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    Imagen
                  </label>
                  <input
                    type="file"
                    accept="image/*"
                    onChange={(e) => setImageFile(e.target.files?.[0] || null)}
                    className="hidden"
                    id="file-upload"
                  />
                  <label
                    htmlFor="file-upload"
                    className="flex items-center justify-center w-full px-3 py-2 border border-gray-300 bg-white text-sm text-gray-700 hover:bg-gray-50 cursor-pointer"
                  >
                    <svg className="w-4 h-4 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16l4.586-4.586a2 2 0 012.828 0L16 16m-2-2l1.586-1.586a2 2 0 012.828 0L20 14m-6-6h.01M6 20h12a2 2 0 002-2V6a2 2 0 00-2-2H6a2 2 0 00-2 2v12a2 2 0 002 2z" />
                    </svg>
                    {imageFile ? imageFile.name : 'Seleccionar archivo'}
                  </label>
                </div>
              </div>

              <div className="flex gap-2">
                <button
                  type="submit"
                  disabled={loading}
                  className="px-4 py-2 bg-blue-600 text-white text-sm font-medium hover:bg-blue-700 disabled:opacity-50"
                >
                  {loading ? 'Guardando...' : 'Guardar'}
                </button>
                <button
                  type="button"
                  onClick={() => setShowForm(false)}
                  className="px-4 py-2 border border-gray-300 text-gray-700 text-sm font-medium hover:bg-gray-50"
                >
                  Cancelar
                </button>
              </div>
            </form>
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
