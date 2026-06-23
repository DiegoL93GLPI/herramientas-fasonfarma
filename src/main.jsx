import React, { useState } from 'react'
import ReactDOM from 'react-dom/client'
import TarifaMOD from './TarifaMOD.jsx'
import EstimadorCostos from './EstimadorCostos.jsx'

function App() {
  const [herramienta, setHerramienta] = useState(null)

  if (herramienta === 'mod') return <><NavBar onBack={() => setHerramienta(null)} titulo="Calculador de Tarifa MOD" /><TarifaMOD /></>
  if (herramienta === 'costos') return <><NavBar onBack={() => setHerramienta(null)} titulo="Estimador de Costos por Orden" /><EstimadorCostos /></>

  return (
    <div style={{
      minHeight: '100vh', background: '#0c0f18', color: '#dde1ea',
      fontFamily: "'Inter','Segoe UI',system-ui,sans-serif",
      display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '2rem 1rem'
    }}>
      <div style={{ maxWidth: 480, width: '100%' }}>
        <p style={{ fontSize: '.7rem', letterSpacing: '.18em', textTransform: 'uppercase', color: '#4a7fa5', marginBottom: '.5rem', textAlign: 'center' }}>
          FasonFarma
        </p>
        <h1 style={{ fontSize: '1.75rem', fontWeight: 800, letterSpacing: '-.03em', color: '#dde1ea', textAlign: 'center', marginBottom: '.4rem' }}>
          Herramientas de <span style={{ color: '#5ba3f5' }}>costos</span>
        </h1>
        <p style={{ fontSize: '.82rem', color: '#4b5568', textAlign: 'center', marginBottom: '2.5rem' }}>
          Produccion farmaceutica por contrato
        </p>

        <div style={{ display: 'flex', flexDirection: 'column', gap: '1rem' }}>
          <button onClick={() => setHerramienta('mod')} style={cardStyle}>
            <div style={cardNumStyle}>1</div>
            <div>
              <div style={cardTitleStyle}>Calculador de Tarifa MOD</div>
              <div style={cardDescStyle}>Salario + IPS + Aguinaldo + Horas extra + Otros costos. Exporta a Excel.</div>
            </div>
          </button>
          <button onClick={() => setHerramienta('costos')} style={cardStyle}>
            <div style={cardNumStyle}>2</div>
            <div>
              <div style={cardTitleStyle}>Estimador de Costos por Orden</div>
              <div style={cardDescStyle}>MOD + CIF · Multiples procesos en cascada. Exporta a Excel y PDF.</div>
            </div>
          </button>
        </div>
      </div>
    </div>
  )
}

function NavBar({ onBack, titulo }) {
  return (
    <div style={{
      background: '#0f1420', borderBottom: '1px solid #1e2535',
      padding: '.75rem 1.25rem', display: 'flex', alignItems: 'center', gap: '1rem',
      position: 'sticky', top: 0, zIndex: 100,
    }}>
      <button onClick={onBack} style={{
        background: 'transparent', border: '1px solid #1e2535', borderRadius: '6px',
        color: '#6b7280', fontSize: '.78rem', padding: '.3rem .7rem', cursor: 'pointer',
      }}>← Volver</button>
      <span style={{ fontSize: '.78rem', color: '#9ca3af', fontWeight: 600 }}>{titulo}</span>
    </div>
  )
}

const cardStyle = {
  background: '#141926', border: '1px solid #1e2535', borderRadius: '12px',
  padding: '1.25rem 1.25rem', display: 'flex', alignItems: 'flex-start', gap: '1rem',
  cursor: 'pointer', textAlign: 'left', transition: 'border-color .15s',
  color: 'inherit',
}
const cardNumStyle = {
  width: 32, height: 32, borderRadius: '50%', background: '#172038',
  border: '1px solid #2a3f5f', display: 'flex', alignItems: 'center', justifyContent: 'center',
  fontSize: '.8rem', fontWeight: 700, color: '#5ba3f5', flexShrink: 0,
}
const cardTitleStyle = { fontSize: '.95rem', fontWeight: 700, color: '#dde1ea', marginBottom: '.3rem' }
const cardDescStyle = { fontSize: '.78rem', color: '#4b5568', lineHeight: 1.5 }

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode><App /></React.StrictMode>
)
