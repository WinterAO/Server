# WinterAO Resurrection - Servidor

![Visual Basic](https://img.shields.io/badge/Visual%20Basic-6.0-blue)

## 📋 Descripción

Este repositorio contiene el código fuente del servidor WinterAO Resurrection, un mod del clásico MMORPG Argentum Online. El servidor está construido en Visual Basic 6.0 y maneja todos los aspectos del juego multijugador.

## ✨ Características Principales

### Sistemas Core
- **Motor de red TCP**: Gestión eficiente de múltiples conexiones simultáneas
- **Sistema de cuentas y personajes**: Gestión completa de usuarios y sus personajes
- **Base de datos integrada**: Persistencia de datos de jugadores, NPCs, items y mapas
- **Sistema de seguridad**: SHA256, detección de anti-cheat y anti-mass clone
- **Sistema de logging**: Registro completo de eventos y acciones del servidor

### Sistemas de Juego
- **IA de NPCs**: Sistema de inteligencia artificial para enemigos y NPCs
- **Sistema de combate completo**: PvP, PvE y cálculo de daño
- **Sistema de hechizos**: Implementación completa de magias y skills
- **Sistema de trabajos**: Herrero, carpintero, minería, pesca, etc.
- **Sistema de clanes (Guilds)**: Gestión de clanes con jerarquías y guerras
- **Sistema de party**: Grupos de jugadores con distribución de experiencia
- **Sistema de comercio**: Comercio entre jugadores y con NPCs
- **Sistema de quests**: Misiones y sistema de recompensas
- **Sistema de facciones**: Ejército Real vs Legión Oscura

### Características Avanzadas
- **PathFinding**: Navegación inteligente de NPCs
- **Sistema de áreas**: Optimización de red y procesamiento
- **Sistema de duelos**: PvP 1v1 y ranked
- **Eventos de mapas**: Eventos dinámicos del servidor
- **Sistema de fortalezas**: Conquista de territorios por clanes
- **Arena Rinkel**: Sistema de torneos automatizado
- **Sistema de subastas**: Mercado de items entre jugadores
- **Sistema de retos**: Desafíos y competiciones
- **Clima dinámico**: Sistema de clima variable por mapa

### Herramientas de Administración
- **Panel de administración**: Interfaz GUI para gestión del servidor
- **Consola de comandos**: Sistema completo de comandos GM/Admin
- **Sistema de estadísticas**: Monitoreo en tiempo real del servidor
- **Herramienta ao-easy-stats**: Generación de estadísticas del juego
- **Sistema de logs**: Registros detallados de todas las acciones

## 🛠️ Requisitos Técnicos

### Para compilar y ejecutar:
- **Visual Basic 6.0** (IDE completo)
- **Windows XP o superior** (recomendado Windows 7/10 con modo compatibilidad)
- **MySQL**

## 📦 Estructura del Proyecto

```
Server/
├── Codigo/                    # Código fuente principal
│   ├── AI_NPC.bas            # Inteligencia artificial de NPCs
│   ├── Protocol.bas          # Protocolo de comunicación (602 KB)
│   ├── Protocol_Write.bas    # Envío de paquetes a clientes
│   ├── SistemaCombate.bas    # Sistema completo de combate
│   ├── modHechizos.bas       # Sistema de hechizos y magias
│   ├── Trabajo.bas           # Sistema de trabajos/oficios
│   ├── modGuilds.bas         # Sistema de clanes
│   ├── TCP.bas               # Gestión de red TCP
│   ├── FileIO.bas            # Entrada/salida de archivos
│   ├── modDatabase.bas       # Gestión de base de datos
│   └── ...                   # +100 módulos más
├── Configuracion/            # Archivos de configuración
├── Fixtures/                 # Datos de prueba
├── stats/                    # Estadísticas del servidor
├── SERVER.VBP                # Proyecto de Visual Basic
└── LICENSE                   # Licencia AGPL-3.0
```

## 🔧 Componentes Principales

### Networking
- `TCP.bas` - Gestión principal de conexiones TCP
- `wskapiAO.bas` - API de Winsock personalizada
- `wsksock.bas` - Implementación de sockets
- `SecurityIp.bas` - Seguridad y baneos por IP

### Protocolo
- `Protocol.bas` - Manejo de paquetes cliente-servidor (archivo principal 602 KB)
- `Protocol_Write.bas` - Funciones de envío de datos a clientes
- `modSendData.bas` - Funciones de broadcasting

### Lógica de Juego
- `GameLogic.bas` - Lógica principal del juego
- `mMainLoop.bas` - Loop principal del servidor
- `SistemaCombate.bas` - Sistema de combate completo
- `AI_NPC.bas` - IA de enemigos y criaturas
- `modHechizos.bas` - Sistema de hechizos y efectos
- `Trabajo.bas` - Sistema de profesiones

### Gestión de Usuarios
- `Modulo_UsUaRiOs.bas` - Gestión de usuarios (108 KB)
- `Cuentas.bas` - Sistema de cuentas
- `Characters.bas` - Gestión de personajes
- `InvUsuario.bas` - Inventario de jugadores

### Sistemas Sociales
- `modGuilds.bas` - Sistema de clanes/guilds
- `clsClan.cls` - Clase de clan
- `clsClanPretoriano.cls` - Clanes pretorianos (97 KB)
- `mdParty.bas` - Sistema de party/grupo
- `clsParty.cls` - Clase de party
- `Amigos.bas` - Sistema de amigos

### Base de Datos
- `FileIO.bas` - E/S de archivos de datos
- `modDatabase.bas` - Operaciones de base de datos
- `clsDataBase.cls` - Clase de base de datos
- `clsIniManager.cls` - Gestión de archivos INI
- `clsIniReader.cls` - Lectura de configuración

### NPCs y Criaturas
- `AI_NPC.bas` - Inteligencia artificial de NPCs
- `MODULO_NPCs.bas` - Gestión de NPCs
- `PathFinding.bas` - Navegación de NPCs
- `ModInvocaciones.bas` - Sistema de invocaciones

### Combate y Facciones
- `SistemaCombate.bas` - Sistema de combate
- `ModFacciones.bas` - Sistema de facciones
- `praetorians.bas` - Pretorianos
- `modArenaRinkel.bas` - Arena de combate
- `Retos.bas` - Sistema de retos/desafíos
- `ModDuelosClasicos.bas` - Duelos clásicos
- `modDuelosRanked.bas` - Duelos ranked

### Comercio y Economía
- `Comercio.bas` - Comercio con NPCs
- `mdlCOmercioConUsuario.bas` - Comercio entre jugadores
- `modBanco.bas` - Sistema bancario
- `modSubastas.bas` - Sistema de subastas

### Quests y Eventos
- `modQuests.bas` - Sistema de misiones
- `clsEventoMapa.cls` - Eventos de mapa
- `modFortalezas.bas` - Sistema de fortalezas
- `clsFortalezas.cls` - Clase de fortalezas
- `Mod_ClanvsClan.bas` - Guerras de clanes

### Administración
- `Admin.bas` - Funciones administrativas
- `modAntiCheat.bas` - Sistema anti-trampas
- `clsAntiMassClon.cls` - Prevención de clonación masiva
- `clsSecurity.cls` - Seguridad general
- `CSHA256.cls` - Hashing SHA256
- `Logs.bas` - Sistema de logs

### Utilidades
- `General.bas` - Funciones generales
- `Matematicas.bas` - Funciones matemáticas
- `Statistics.bas` - Estadísticas del servidor
- `Queue.bas` - Estructuras de cola
- `ModAreas.bas` - Sistema de áreas del mapa
- `modClimas.bas` - Sistema de clima

### Interfaces de Usuario
- `frmMain.frm` - Ventana principal del servidor
- `frmServidor.frm` - Panel del servidor
- `frmAdmin.frm` - Panel administrativo
- `frmUserList.frm` - Lista de usuarios conectados
- `frmTrafic.frm` - Monitor de tráfico
- `FrmInterv.frm` - Interfaz de intervención
- `FrmStat.frm` - Estadísticas

## 🚀 Compilación

1. Abre el archivo `SERVER.VBP` con Visual Basic 6.0
2. Asegúrate de tener todas las dependencias instaladas
3. Configura los archivos en la carpeta `Configuracion/`
4. Compila el proyecto

## ⚙️ Configuración

### Archivos de Configuración (carpeta `Configuracion/`)
- Configuración del servidor (IP, puerto, límites)
- Configuración de NPCs y criaturas
- Configuración de items y objetos
- Configuración de hechizos y skills
- Configuración de mapas
- Balanceo de combate y experiencia
- Configuración de clanes y facciones

### Base de Datos
El servidor utiliza base de datos MySQL para cuentas y personajes.

## 📊 Monitoreo

### Herramienta ao-easy-stats.exe
Incluye una herramienta para generar estadísticas del servidor:
- Jugadores conectados
- Actividad del servidor
- Estadísticas de combate
- Economía del juego

## 🔒 Seguridad

El servidor implementa múltiples capas de seguridad:
- Hashing SHA256 para contraseñas
- Sistema de detección de anti-cheat
- Protección contra mass-cloning
- Sistema de baneos por IP
- Validación de paquetes
- Protección contra flooding
## 🔗 Enlaces Relacionados

- [Repositorio del Servidor](https://github.com/WinterAO/Server)
- [Herramientas y recursos](https://github.com/orgs/WinterAO/repositories)

## 🐛 Problemas Conocidos

- Compatibilidad limitada con Windows 10/11 (requiere modo compatibilidad)
- Algunas funciones pueden requerir permisos de administrador
- El servidor legacy puede necesitar configuraciones específicas de firewall

## 🤝 Contribuir

Las contribuciones son bienvenidas. Por favor:
1. Fork el proyecto
2. Crea una rama para tu feature (`git checkout -b feature/AmazingFeature`)
3. Commit tus cambios (`git commit -m 'Add some AmazingFeature'`)
4. Push a la rama (`git push origin feature/AmazingFeature`)
5. Abre un Pull Request

## ⭐ Agradecimientos

Este proyecto es un mod de **Argentum Online**. Todo el crédito original corresponde a los creadores de Argentum Online y sus contribuidores.
