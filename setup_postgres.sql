-- ============================================================
--  CUADRA ERRE - Script de base de datos para PostgreSQL
--  Ejecutar una sola vez al configurar el servidor
-- ============================================================

CREATE SEQUENCE seq_roles;
CREATE SEQUENCE seq_usuarios;
CREATE SEQUENCE seq_pacientes;
CREATE SEQUENCE seq_familiares;
CREATE SEQUENCE seq_equinos;
CREATE SEQUENCE seq_areas;
CREATE SEQUENCE seq_sesiones;
CREATE SEQUENCE seq_registro_equino;
CREATE SEQUENCE seq_notificaciones;
CREATE SEQUENCE seq_auditoria;
CREATE SEQUENCE seq_configuracion;
CREATE SEQUENCE seq_mantenimiento;

CREATE TABLE roles (
  id          INTEGER      NOT NULL DEFAULT nextval('seq_roles'),
  nombre      VARCHAR(50)  NOT NULL,
  descripcion VARCHAR(200),
  CONSTRAINT pk_roles PRIMARY KEY (id)
);

CREATE TABLE usuarios (
  id             INTEGER       NOT NULL DEFAULT nextval('seq_usuarios'),
  username       VARCHAR(80)   NOT NULL,
  password_hash  VARCHAR(255)  NOT NULL,
  nombre         VARCHAR(100)  NOT NULL,
  apellido       VARCHAR(100),
  email          VARCHAR(150),
  telefono       VARCHAR(20),
  rol_id         INTEGER       NOT NULL,
  activo         SMALLINT      DEFAULT 1,
  fecha_creacion TIMESTAMP     DEFAULT CURRENT_TIMESTAMP,
  CONSTRAINT pk_usuarios PRIMARY KEY (id),
  CONSTRAINT uq_username UNIQUE (username),
  CONSTRAINT fk_usu_rol FOREIGN KEY (rol_id) REFERENCES roles(id)
);

CREATE TABLE pacientes (
  id                    INTEGER       NOT NULL DEFAULT nextval('seq_pacientes'),
  nombre                VARCHAR(100)  NOT NULL,
  apellido              VARCHAR(100)  NOT NULL,
  fecha_nacimiento      DATE,
  diagnostico           VARCHAR(500),
  alergias              VARCHAR(300),
  medicamentos          VARCHAR(300),
  contacto_emergencia   VARCHAR(150),
  telefono_emergencia   VARCHAR(20),
  email_familiar        VARCHAR(150),
  activo                SMALLINT      DEFAULT 1,
  fecha_registro        TIMESTAMP     DEFAULT CURRENT_TIMESTAMP,
  CONSTRAINT pk_pacientes PRIMARY KEY (id)
);

CREATE TABLE familiares (
  id          INTEGER       NOT NULL DEFAULT nextval('seq_familiares'),
  paciente_id INTEGER       NOT NULL,
  usuario_id  INTEGER,
  nombre      VARCHAR(100)  NOT NULL,
  apellido    VARCHAR(100)  NOT NULL,
  parentesco  VARCHAR(50),
  telefono    VARCHAR(20),
  email       VARCHAR(150),
  CONSTRAINT pk_familiares PRIMARY KEY (id),
  CONSTRAINT fk_fam_paciente FOREIGN KEY (paciente_id) REFERENCES pacientes(id)
);

CREATE TABLE equinos (
  id              INTEGER      NOT NULL DEFAULT nextval('seq_equinos'),
  nombre          VARCHAR(80)  NOT NULL,
  raza            VARCHAR(80),
  edad            INTEGER,
  color           VARCHAR(50),
  estado          VARCHAR(30)  DEFAULT 'DISPONIBLE',
  notas_salud     VARCHAR(500),
  ultima_revision DATE,
  activo          SMALLINT     DEFAULT 1,
  CONSTRAINT pk_equinos PRIMARY KEY (id)
);

CREATE TABLE areas (
  id          INTEGER       NOT NULL DEFAULT nextval('seq_areas'),
  nombre      VARCHAR(100)  NOT NULL,
  descripcion VARCHAR(200),
  capacidad   INTEGER       DEFAULT 1,
  activo      SMALLINT      DEFAULT 1,
  CONSTRAINT pk_areas PRIMARY KEY (id)
);

CREATE TABLE sesiones (
  id                    INTEGER       NOT NULL DEFAULT nextval('seq_sesiones'),
  paciente_id           INTEGER       NOT NULL,
  terapeuta_id          INTEGER       NOT NULL,
  equino_id             INTEGER,
  area_id               INTEGER,
  fecha_hora            TIMESTAMP     NOT NULL,
  duracion_min          INTEGER       DEFAULT 45,
  estado                VARCHAR(30)   DEFAULT 'PROGRAMADA',
  objetivos             VARCHAR(500),
  notas_sesion          VARCHAR(1000),
  recomendaciones_casa  VARCHAR(500),
  estado_paciente       VARCHAR(50),
  area_trabajada        VARCHAR(80),
  email_enviado         SMALLINT      DEFAULT 0,
  recordatorio_enviado  SMALLINT      DEFAULT 0,
  fecha_registro        TIMESTAMP     DEFAULT CURRENT_TIMESTAMP,
  CONSTRAINT pk_sesiones PRIMARY KEY (id),
  CONSTRAINT fk_ses_paciente  FOREIGN KEY (paciente_id)  REFERENCES pacientes(id),
  CONSTRAINT fk_ses_terapeuta FOREIGN KEY (terapeuta_id) REFERENCES usuarios(id),
  CONSTRAINT fk_ses_equino    FOREIGN KEY (equino_id)    REFERENCES equinos(id),
  CONSTRAINT fk_ses_area      FOREIGN KEY (area_id)      REFERENCES areas(id)
);

CREATE TABLE registro_equino (
  id             INTEGER       NOT NULL DEFAULT nextval('seq_registro_equino'),
  equino_id      INTEGER       NOT NULL,
  fecha          DATE          NOT NULL,
  encargado_id   INTEGER,
  estado_fisico  VARCHAR(50),
  estado_animo   VARCHAR(50),
  disponible     SMALLINT      DEFAULT 1,
  observaciones  VARCHAR(500),
  CONSTRAINT pk_reg_eq     PRIMARY KEY (id),
  CONSTRAINT fk_reg_equino FOREIGN KEY (equino_id) REFERENCES equinos(id)
);

CREATE TABLE notificaciones (
  id             INTEGER        NOT NULL DEFAULT nextval('seq_notificaciones'),
  sesion_id      INTEGER,
  familiar_id    INTEGER,
  tipo           VARCHAR(50),
  mensaje        VARCHAR(1000),
  enviada        SMALLINT       DEFAULT 0,
  email_destino  VARCHAR(150),
  fecha_envio    TIMESTAMP,
  fecha_creacion TIMESTAMP      DEFAULT CURRENT_TIMESTAMP,
  CONSTRAINT pk_notif      PRIMARY KEY (id),
  CONSTRAINT fk_not_sesion FOREIGN KEY (sesion_id)   REFERENCES sesiones(id),
  CONSTRAINT fk_not_fam    FOREIGN KEY (familiar_id) REFERENCES familiares(id)
);

CREATE TABLE auditoria (
  id             INTEGER       NOT NULL DEFAULT nextval('seq_auditoria'),
  usuario_id     INTEGER,
  accion         VARCHAR(100),
  tabla_afectada VARCHAR(50),
  registro_id    INTEGER,
  detalle        VARCHAR(500),
  fecha          TIMESTAMP     DEFAULT CURRENT_TIMESTAMP,
  CONSTRAINT pk_auditoria PRIMARY KEY (id)
);

CREATE TABLE configuracion (
  id                 INTEGER       NOT NULL DEFAULT nextval('seq_configuracion'),
  nombre_centro      VARCHAR(150)  DEFAULT 'Cuadra Erre',
  telefono           VARCHAR(20),
  direccion          VARCHAR(300),
  email_contacto     VARCHAR(150),
  color_primario     VARCHAR(10)   DEFAULT '#1e3a5f',
  recordatorio_24h   SMALLINT      DEFAULT 1,
  horas_anticipacion INTEGER       DEFAULT 24,
  CONSTRAINT pk_config PRIMARY KEY (id)
);

CREATE TABLE mantenimiento_equino (
  id               INTEGER       NOT NULL DEFAULT nextval('seq_mantenimiento'),
  equino_id        INTEGER       NOT NULL,
  tipo             VARCHAR(50)   NOT NULL,
  fecha_programada DATE          NOT NULL,
  fecha_realizada  DATE,
  estado           VARCHAR(30)   DEFAULT 'PENDIENTE',
  notas            VARCHAR(500),
  registrado_por   INTEGER,
  fecha_creacion   TIMESTAMP     DEFAULT CURRENT_TIMESTAMP,
  CONSTRAINT pk_mant_eq    PRIMARY KEY (id),
  CONSTRAINT fk_mant_equino  FOREIGN KEY (equino_id)      REFERENCES equinos(id),
  CONSTRAINT fk_mant_usuario FOREIGN KEY (registrado_por) REFERENCES usuarios(id)
);

CREATE INDEX idx_sesiones_fecha ON sesiones (fecha_hora);
CREATE INDEX idx_sesiones_pac   ON sesiones (paciente_id);
CREATE INDEX idx_pacientes_nom  ON pacientes (apellido, nombre);
CREATE INDEX idx_equinos_estado ON equinos (estado);
CREATE INDEX idx_mant_fecha     ON mantenimiento_equino (fecha_programada);
CREATE INDEX idx_mant_equino    ON mantenimiento_equino (equino_id);

-- ============================================================
--  DATOS INICIALES
-- ============================================================

INSERT INTO roles (nombre, descripcion) VALUES
  ('ADMINISTRADOR',    'Acceso total al sistema'),
  ('TERAPEUTA',        'Gestion de sesiones y expedientes'),
  ('COORDINADOR',      'Agenda y reportes'),
  ('ENCARGADO_EQUINOS','Registro de caballos');

INSERT INTO areas (nombre, descripcion, capacidad) VALUES
  ('Picadero 1',       'Pista principal',    1),
  ('Picadero 2',       'Pista secundaria',   1),
  ('Sala de Evaluacion','Evaluacion inicial', 2);

INSERT INTO equinos (nombre, raza, edad, color, estado) VALUES
  ('Tornado',  'Cuarto de Milla', 8,  'Bayo',    'DISPONIBLE'),
  ('Luna',     'Andaluz',         6,  'Blanco',  'DISPONIBLE'),
  ('Rayo',     'Mestizo',         10, 'Negro',   'DISPONIBLE'),
  ('Estrella', 'Pura Sangre',     7,  'Alazan',  'DISPONIBLE'),
  ('Oso',      'Azteca',          5,  'Castano', 'DISPONIBLE');

INSERT INTO pacientes (nombre, apellido, fecha_nacimiento, diagnostico) VALUES
  ('Sofia',  'Garcia', '2015-03-12', 'Trastorno del Espectro Autista'),
  ('Miguel', 'Torres', '2012-07-08', 'TDAH'),
  ('Lucas',  'Perez',  '2014-05-30', 'Sindrome de Down');

-- Usuario admin (contrasena: #Adm1234.)
INSERT INTO usuarios (username, password_hash, nombre, apellido, email, rol_id) VALUES
  ('admin', 'scrypt:32768:8:1$pFuZ1dhdz5i7UpGP$77541f5cffb6569ffdddf616ba6cd6971492422deb796342923af37e2a8a4fd95a41409c9baaf50abb83d74b03652f57f167611e9b69e772c3153112a2110d48', 'Administrador', 'Sistema', 'admin@cuadraerre.mx', 1);

INSERT INTO configuracion (nombre_centro, telefono, direccion, email_contacto, color_primario, recordatorio_24h, horas_anticipacion)
VALUES ('Cuadra Erre', '624-000-0000', 'Los Cabos, Baja California Sur', 'cuadraerreoficial@gmail.com', '#1e3a5f', 1, 24);
