import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import './index.css'
import ValuationApp from './ValuationApp.jsx'

createRoot(document.getElementById('root')).render(
  <StrictMode>
    <ValuationApp />
  </StrictMode>,
)
