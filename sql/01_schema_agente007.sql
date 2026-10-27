-- ==============================================================================
-- PROYECTO: AGENTE 007 (Licencia para Resolver)
-- SCRIPT 01: Arquitectura de Base de Datos para Trouble Tickets y Cruces
-- MOTOR: PostgreSQL (Supabase)
-- ==============================================================================

-- 1. Habilitar extensiones necesarias
CREATE EXTENSION IF NOT EXISTS "uuid-ossp";

-- 2. Limpieza previa (por si se reejecuta)
DROP TABLE IF EXISTS ticket_audit_logs CASCADE;
DROP TABLE IF EXISTS tickets CASCADE;
DROP TABLE IF EXISTS customers CASCADE;
DROP TABLE IF EXISTS ticket_catalog CASCADE;
DROP TYPE IF EXISTS ticket_type_enum CASCADE;
DROP TYPE IF EXISTS ticket_status_enum CASCADE;

-- 3. Tipos enumerados estrictos (Garantiza integridad de datos contra errores humanos)
CREATE TYPE ticket_type_enum AS ENUM (
    'INCIDENTE',     -- Algo falló o se interrumpió
    'REQUERIMIENTO'  -- Solicitud de servicio o trámite (ej. Activación)
);

CREATE TYPE ticket_status_enum AS ENUM (
    'PENDIENTE',         -- Recién llegado (esperando análisis)
    'EN_PROCESO_007',    -- Agente 007 está ejecutando los cruces de datos
    'AUTO_RESUELTO',     -- Resuelto de forma 100% autónoma por Agente 007
    'ESCALADO_HUMANO',   -- Requiere intervención o aprobación de un humano (HITL)
    'CERRADO'            -- Finalizado
);

-- 4. Catálogo de Tipificaciones y Acciones (El reemplazo inteligente del Excel)
CREATE TABLE ticket_catalog (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    tipo ticket_type_enum NOT NULL,
    categoria VARCHAR(100) NOT NULL,          -- Ej. 'ACTIVACION', 'CONECTIVIDAD', 'FACTURACION'
    accion_sugerida VARCHAR(150) NOT NULL,    -- Ej. 'REALIZAR_ACTIVACION', 'RESET_PUERTO', 'VALIDAR_PAGO'
    es_automatizable BOOLEAN DEFAULT FALSE,   -- Si TRUE, Agente 007 tiene permiso de resolverlo solo
    descripcion_accion TEXT,
    created_at TIMESTAMPTZ DEFAULT NOW()
);

-- 5. Tabla de Clientes / Servicios (Para que Agente 007 haga los "Cruces" automáticos)
CREATE TABLE customers (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    cuenta_cliente VARCHAR(50) UNIQUE NOT NULL, -- Número de cuenta o contrato
    nombre_completo VARCHAR(150) NOT NULL,
    telefono VARCHAR(30) NOT NULL,
    plan_servicio VARCHAR(100) NOT NULL,
    estado_servicio VARCHAR(50) DEFAULT 'INACTIVO', -- 'ACTIVO', 'INACTIVO', 'SUSPENDIDO'
    saldo_pendiente NUMERIC(10, 2) DEFAULT 0.00,
    created_at TIMESTAMPTZ DEFAULT NOW()
);

-- 6. Tabla Central de Tickets (Trouble Tickets)
CREATE TABLE tickets (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    folio SERIAL UNIQUE,                         -- Folio numérico amigable (ej. TT-1001)
    cuenta_cliente VARCHAR(50) REFERENCES customers(cuenta_cliente) ON DELETE SET NULL,
    
    -- Ingesta inicial
    canal_origen VARCHAR(50) DEFAULT 'WHATSAPP', -- 'WHATSAPP', 'EXCEL_BATCH', 'WEB'
    descripcion_original TEXT NOT NULL,         -- Lo que escribió el técnico o cliente
    
    -- Tipificación clasificada por Agente 007
    tipo ticket_type_enum DEFAULT 'REQUERIMIENTO',
    categoria VARCHAR(100),                      -- Extraído por IA
    accion VARCHAR(150),                         -- Extraído por IA (ej. 'REALIZAR_ACTIVACION')
    
    -- Estado y Resolución
    status ticket_status_enum DEFAULT 'PENDIENTE',
    confianza_ia NUMERIC(3, 2),                  -- De 0.00 a 1.00 (métrica de certeza del agente)
    comentario_resolucion TEXT,                 -- El comentario técnico generado por Agente 007
    resolucion_automatica BOOLEAN DEFAULT FALSE,
    
    -- Métricas de SLA (Tiempos de respuesta)
    created_at TIMESTAMPTZ DEFAULT NOW(),
    updated_at TIMESTAMPTZ DEFAULT NOW(),
    resolved_at TIMESTAMPTZ
);

-- 7. Tabla de Auditoría y Trazabilidad (Crítica para Ciberseguridad y evaluaciones de Frontier)
CREATE TABLE ticket_audit_logs (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    ticket_id UUID REFERENCES tickets(id) ON DELETE CASCADE,
    nodo_ejecutor VARCHAR(100) NOT NULL,         -- 'SanitizerNode', 'CrossReferencerNode', 'ResolutionNode'
    accion_ejecutada VARCHAR(150) NOT NULL,
    detalles JSONB,                              -- Datos exactos del cruce realizado
    created_at TIMESTAMPTZ DEFAULT NOW()
);

-- 8. Índices de Alto Rendimiento para Cruces Inmediatos (Búsqueda en milisegundos)
CREATE INDEX idx_tickets_cuenta_cliente ON tickets(cuenta_cliente);
CREATE INDEX idx_tickets_status ON tickets(status);
CREATE INDEX idx_customers_cuenta ON customers(cuenta_cliente);

-- 9. Insertar Datos de Prueba (Semilla inicial para probar de inmediato)

-- Insertamos el catálogo con tu caso específico: 'Activación - Realizar activación'
INSERT INTO ticket_catalog (tipo, categoria, accion_sugerida, es_automatizable, descripcion_accion) VALUES
('REQUERIMIENTO', 'ACTIVACION', 'REALIZAR_ACTIVACION', TRUE, 'Activar el servicio si el cliente no tiene saldo pendiente'),
('REQUERIMIENTO', 'CAMBIO_PLAN', 'APLICAR_CAMBIO_PLAN', TRUE, 'Actualizar el plan contratado en el sistema'),
('INCIDENTE', 'CONECTIVIDAD', 'DIAGNOSTICO_LINEA', FALSE, 'Falla física o corte de fibra. Requiere cuadrilla en campo'),
('INCIDENTE', 'FACTURACION', 'ACLARACION_CARGO', FALSE, 'Disputa de cobro mayor a $100. Requiere revisión humana');

-- Insertamos 2 clientes de prueba para los cruces:
INSERT INTO customers (cuenta_cliente, nombre_completo, telefono, plan_servicio, estado_servicio, saldo_pendiente) VALUES
('CLI-1001', 'Juan Pérez Morales', '+525512345678', 'Plan Fibra 200MB', 'INACTIVO', 0.00),
('CLI-1002', 'María López Santos', '+525587654321', 'Plan Negocio 500MB', 'SUSPENDIDO', 450.00);

-- Habilitar Supabase Realtime para la tabla de tickets (Permite que el Dashboard se mueva solo)
ALTER PUBLICATION supabase_realtime ADD TABLE tickets;
