-- phpMyAdmin SQL Dump
-- version 5.0.4deb2+deb11u1
-- https://www.phpmyadmin.net/
--
-- Servidor: localhost:3306
-- Tiempo de generación: 29-09-2022 a las 21:21:51
-- Versión del servidor: 10.5.15-MariaDB-0+deb11u1
-- Versión de PHP: 7.4.30

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Base de datos: `winterweb`
--
CREATE DATABASE IF NOT EXISTS `winterweb` DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci;
USE `winterweb`;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `cuentas`
--

CREATE TABLE `cuentas` (
  `id` int(20) NOT NULL,
  `username` varchar(50) DEFAULT NULL,
  `email` varchar(50) DEFAULT NULL,
  `password` varchar(250) DEFAULT NULL,
  `salt` varchar(250) DEFAULT NULL,
  `id_recuperacion` varchar(250) DEFAULT NULL,
  `id_confirmacion` varchar(250) DEFAULT NULL,
  `last_ip` varchar(32) DEFAULT NULL,
  `gemas` int(12) DEFAULT 0,
  `status` int(1) DEFAULT 0,
  `macaddress` varchar(128) DEFAULT NULL,
  `serialhd` varchar(128) DEFAULT NULL,
  `vip` date DEFAULT '1001-01-01',
  `role` enum('user','consejero','semidios','dios','admin') DEFAULT 'user',
  `createdAt` datetime DEFAULT NULL,
  `updatedAt` datetime DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

--
-- Volcado de datos para la tabla `cuentas`
--

INSERT INTO `cuentas` (`id`, `username`, `email`, `password`, `salt`, `id_recuperacion`, `id_confirmacion`, `last_ip`, `gemas`, `status`, `macaddress`, `serialhd`, `vip`, `role`, `createdAt`, `updatedAt`) VALUES
(1, 'WinterStaff', 'manueljsandalio@gmail.com', '332643853d449f9c88af3fd632056bcaff4d826ba206adbfcc15f4cbd26bea52', ' aaShD3Z2tLouPknCUguvIxIqE2AqIMvB', NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 'admin', '2022-06-23 18:31:44', '2022-06-23 18:31:44'),
(2, 'Lorwik', 'lorwik@gmail.com', 'adc9928f5e46885d6f4aa68144aad2c4890617eb1c42c39a5057d60db071aa01', ' vArdbPKnlvgUHDDSX0ScBTgoSKl1NUGq', NULL, NULL, '127.0.0.1', 0, 1, '4C:CC:6A:B3:F3:5E', '1828424595', '1001-01-01', 'admin', '2022-06-23 18:32:25', '2022-09-29 13:06:34'),
(3, 'Howell', 'Gonzalomir96@gmail.com', '20302ed6262128747f9470d3254a2851818645c9e797c6f055289a6cf750b2b7', ' 1sRfSMn3shayivjpvaIgdFYIWwNSP6Ep', NULL, ' baBUSp5RrL5f', '186.128.245.11', 0, 1, 'B4:2E:99:E8:41:D5', '1996694788', '1001-01-01', 'user', '2022-07-16 23:34:47', '2022-07-30 02:26:19'),
(4, 'sensui', 'asdd@asd.com', '0c4c5420df8b7b11c6af6e4c9038a0cd79469c436d3eb5438ed8de89a30c56e6', 'pg0$jQSrLX8vF)p_qq3HMyO-WmJbo~U1', NULL, '648574', '213.94.53.233', 0, 1, '00:1F:3A:57:DE:58', '1097161887', '1001-01-01', 'user', '2022-08-09 20:31:02', '2022-08-09 20:31:58'),
(5, 'betacuenta', 'asdasdad@asdasd.com', '1ded5aefc595f6643b535fcb074edd6b37cc610890fda38c5ce6f0a54e37cda6', 'tGbbi(XXl7uB!BMnijEE@P$M-x~oApA3', NULL, '929741', NULL, 0, 1, NULL, NULL, '1001-01-01', 'user', '2022-09-29 12:24:51', '2022-09-29 12:24:51'),
(6, 'cuentabeta', 'asdasdad@asdasd.com', '1ded5aefc595f6643b535fcb074edd6b37cc610890fda38c5ce6f0a54e37cda6', 'tGbbi(XXl7uB!BMnijEE@P$M-x~oApA3', NULL, '929741', NULL, 0, 1, NULL, NULL, '1001-01-01', 'user', '2022-09-29 12:24:51', '2022-09-29 12:24:51');

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `noticias`
--

CREATE TABLE `noticias` (
  `id` int(20) NOT NULL,
  `titulo` varchar(250) NOT NULL,
  `cuerpo` varchar(8000) NOT NULL,
  `createdAt` date DEFAULT NULL,
  `updatedAt` date DEFAULT NULL,
  `cuentaid` int(20) DEFAULT 1
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

--
-- Volcado de datos para la tabla `noticias`
--

INSERT INTO `noticias` (`id`, `titulo`, `cuerpo`, `createdAt`, `updatedAt`, `cuentaid`) VALUES
(1, '¡Estrenamos Web!', 'Hace 3 meses, parte del antiguo Staff de WinterAO nos reunimos después de 8 años para emprender una nueva versión del proyecto. Si bien es cierto que para algunos la espera se hace larga desde el Staff podemos decir que estamos desarrollando la nueva versión en un tiempo récord. Y es que en 3 meses de desarrollo, hemos logrado hacer lo que en versiones anteriores nos llevo años hacer. \r\n<br/><br/>\r\nEstamos desarrollando esta versión completamente desde cero, pero con la experiencia de las versiones anteriores, y con los conocimientos y herramientas actuales. Estamos poniendo mucho empeño para diseñar y desarrollar todo al mas mínimo detalle, aprendiendo de los errores pasados.\r\n<br/><br/>\r\nEs por lo que nos enorgullece dar este gran paso abriendo la nueva Web y Wiki al publico, que si bien aun no están terminadas al 100%, es una buena señal de que estamos haciendo progresos y que cada día falta menos para esta online.\r\n<br/><br/>\r\nLe damos especialmente las gracias a Jopi por el increible trabajo que ha estado realizando en la web, y que a día de hoy sigue. También aprovechamos, para dar las gracias a toda la comunidad, que día tras día nos apoya y sigue creciendo. Y le recordamos que nos pueden seguir en las redes sociales, y unirse a nuestro canal de Discord para participar en la comunidad y estar al ultima sobre el desarrollo.', '2022-06-23', '2022-06-23', 1),
(2, '¡Grandes Cambios!', 'Como habréis podido notar, estos últimos meses hemos estado algo inactivos en las redes. Tanto yo (Zharkekl) como Lorwik hemos estado ausentes durante una corta temporada. Este proyecto lo llevamos a cabo en nuestro tiempo libre, pero tanto nosotros como el resto del Staff nos sentimos completamente comprometidos con la comunidad.\r\n<br/><br/>\r\nActualmente nos encontramos gestionando además de WinterAO otros proyectos vinculados a este, es por ello que a partir de ahora pasamos a llamarnos COMUNIDAD WINTER. Queremos hacer algo muy grande y para ello necesitamos ampliar nuestro Staff, así pues desde el día de hoy quedan abiertas las postulaciones para formar parte del Staff de Comunidad Winter, si creen que pueden llegar a formar parte del Staff aquí abajo les dejamos un formulario de postulación al mismo.', '2022-06-23', '2022-06-23', 1),
(3, 'Estrenamos diseño web', 'Cada día trabajamos para construir la infraestructura del proyecto Winter, ya que la nueva versión requiere de muchas features externas al propio juego para poder funcionar correctamente, y poder dar la mejor experiencia al usuario.\r\n<br/><br/>\r\nEl caso de la web no es para menos, ya que será uno de los puntos mas importante de la nueva versión, y es por ello que hemos dedicado tiempo para mejorar su diseño, para que sea mas amigable con el usuario y ofrezca mas información. El remodelado no solo ha sido en diseño, también se ha reprogramado todo utilizando las tecnologías mas punteras. Lejos de estas completa, además del diseño, se han actualizado algunas de las secciones, como \'Multimedia\' y \'Staff\'.\r\n<br/><br/>\r\nAprovecho la ocasión para recordarles que seguimos en búsqueda de Staff, entre ellos un programador web para seguir trabajando en todas las características que ofrecerá.\r\n<br/><br/>\r\nFormulario de postulación\r\n<br/><br/>\r\nSin mas, seguimos trabajando en la nueva versión de Winter, y os prometemos que la espera merecerá la pena! Próximamente daremos mucha mas información acerca de lo que traerá la nueva versión.', '2022-06-23', '2022-06-23', 1),
(4, '¡A por la copa Argentum!', 'En nombre de todo el Staff de Comunidad Winter nos complace informaros que nos hemos inscrito en la primera copa Argentum. Santos, Krosty, Moncho, Rey T y Topito son los jugadores que represantara a Winter en la copa. Conozcan mas de ellos en los videos de presentación\n<br/><br/>\nhttps://www.youtube.com/watch?v=ezEF8QLxasY\n<br/><br/>\nhttps://www.youtube.com/watch?v=uh_FRhr22do', '2022-07-05', '2022-07-05', 1),
(5, '¡Volvemos a la carga!', 'Han sido unos meses de inactividad en el proyecto, los miembros del staff, actualmente somos 3 miembros en el staff para levantar un proyecto gigante y hemos estado muy atareados. Es por ello por lo que hemos decidido abrir próximamente un servidor de Agite llamado \"Battlegrounds\". Este servidor se enfocara principalmente en el PVP varios sistemas exclusivos de Winter y un mundo continuo que fomente el PVP a base de objetivos. Insistimos en que somos muy pocos miembros en el staff de desarrollo para tan arduo trabajo, por lo que las postulaciones a staff de desarrollo siguen abiertas. Mapeador, Graficador, Indexador, Webmaster, realmente cualquier area. Asi que si estas interesando en sumarte a un proyecto serio y grande como WinterAO no dudes en contactarnos. Puedes hacerlo atreves de las redes sociales o por discord contactando directamente con Lorwik. Saludos.', '2022-07-05', '2022-07-05', 1);

--
-- Índices para tablas volcadas
--

--
-- Indices de la tabla `cuentas`
--
ALTER TABLE `cuentas`
  ADD PRIMARY KEY (`id`);

--
-- Indices de la tabla `noticias`
--
ALTER TABLE `noticias`
  ADD PRIMARY KEY (`id`),
  ADD KEY `cuentaid` (`cuentaid`);

--
-- AUTO_INCREMENT de las tablas volcadas
--

--
-- AUTO_INCREMENT de la tabla `cuentas`
--
ALTER TABLE `cuentas`
  MODIFY `id` int(20) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=7;

--
-- AUTO_INCREMENT de la tabla `noticias`
--
ALTER TABLE `noticias`
  MODIFY `id` int(20) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=6;

--
-- Restricciones para tablas volcadas
--

--
-- Filtros para la tabla `noticias`
--
ALTER TABLE `noticias`
  ADD CONSTRAINT `noticias_ibfk_1` FOREIGN KEY (`cuentaid`) REFERENCES `cuentas` (`id`);
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
