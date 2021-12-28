-- phpMyAdmin SQL Dump
-- version 4.6.6deb5
-- https://www.phpmyadmin.net/
--
-- Servidor: localhost:3306
-- Tiempo de generación: 05-09-2020 a las 14:27:26
-- Versión del servidor: 10.3.22-MariaDB-0+deb10u1
-- Versión de PHP: 7.3.19-1~deb10u1

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Base de datos: `comunidadwinter`
--
CREATE DATABASE IF NOT EXISTS `comunidadwinter` DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci;
USE `comunidadwinter`;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `cuentas`
--

CREATE TABLE `cuentas` (
  `id` mediumint(8) UNSIGNED NOT NULL,
  `username` varchar(24) NOT NULL,
  `email` varchar(64) NOT NULL,
  `password` varchar(64) NOT NULL,
  `salt` varchar(32) NOT NULL,
  `id_recuperacion` varchar(32) DEFAULT NULL,
  `date_created` timestamp NULL DEFAULT current_timestamp(),
  `last_ip` varchar(16) DEFAULT NULL,
  `date_last_login` timestamp NULL DEFAULT current_timestamp(),
  `gemas` int(12) DEFAULT 0,
  `status` tinyint(1) NOT NULL DEFAULT 0,
  `id_confirmacion` varchar(128) NOT NULL,
  `macaddress` VARCHAR(128) NULL DEFAULT NULL,
  `serialhd` INT(24) NULL DEFAULT '0'
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

--
-- Volcado de datos para la tabla `cuentas`
--

INSERT INTO `cuentas` (`id`, `username`, `email`, `password`, `salt`, `id_recuperacion`, `date_created`, `last_ip`, `date_last_login`, `gemas`, `status`, `id_confirmacion`) VALUES
(2, 'Lorwik', 'lorwik@gmail.com', '330d98ee11071b977ca70f6cb9d680855da8d77f30c312dde101be74775153a5', '03f782fd99e7787c1cac925a452d6264', NULL, '2020-06-04 18:39:29', '79.108.8.202', '2020-10-29 17:41:40', 0, 1, 'yyUxfFxNP7QySeiUC5P6YK328P5sOh8N'),
(8, 'sANTO', 'sonrisa_eventos@hotmail.com', '71f46a44617e6e6f3f72b68b7980787eec16d69f7d61c52d8ade8b860537764f', 'k0UJiREgZqmXEtZiLfefMEiThYiuSoYJ', NULL, '2020-06-04 19:56:08', '186.138.38.130', '2020-10-30 06:44:44', 0, 1, 'VERIFICADA'),
(10, 'Definiun', 'Definiun@gmail.com', '892a6e79a67699642e8330a24a87d72184842c4dcd097bd9b9f4d346fadb7951', 'D8pgoUVJZYC0vz3imklk6ZUYLXS2gPYh', NULL, '2020-06-04 20:22:58', '192.168.1.7', '2020-09-03 20:45:10', 0, 1, 'JkDnxG7PR0j178DSut7DRHMt4UxSymvY'),
(11, 'jopiodz', 'jopiodz00@hotmail.com', '05a679ae416542c12c68e0d4e0098db0d42bc2bf3ca246cb5f3534c21187bf90', '33e7c70a48687adc72b42a7ebf824bb7', 'fCxq7Gs4LHZuvjbkckZO4j6Eq0jcEo43', '2020-06-04 20:41:35', '190.244.223.128', '2020-08-30 18:44:34', 25, 1, 'VERIFICADA'),
(12, 'Howell', 'pjs_Del_ao@hotmail.com', '99b6670de91e04a20dd7abf6093beae193f2745bf865f6fd28a2a5cab25c6c6c', '8M2Y7BZ0JtMLWq3DZXgF3JHEJj7ybVmV', NULL, '2020-06-06 00:04:43', '190.50.68.240', '2020-10-28 22:39:26', 0, 1, 'RfwE5NbuCvKbR6Wzduh5krLFFuTmq8zj'),
(15, 'betatester', 'betatester@gmail.com', '121d0f75ef8334057d85b2e56008e74b4277d7f55728e1da888b6967ed1b2e49', '09347761eec7c3aae5da1f8f8228dec9', NULL, '2020-10-28 18:29:29', '79.108.8.202', '2020-10-28 18:29:29', 0, 0, '33b08bca28c9697b'),
(16, 'Portu', 'portutf@hotmail.com', '743830216c05f309cfe68ad490d897dc6ed0e0e7af5a0b7472816b986ed63eb8', '2696d42a0cdec90d376706c50fe2b2a2', NULL, '2020-10-28 18:32:31', '79.108.8.202', '2020-10-28 18:32:31', 0, 0, 'aed2fccc84650be7'),
(17, 'Sinon', 'francononis24@gmail.com', '857a5c601fa513c505ecf293ab33e1ab3922adb66f70d862997a50dc554de1c1', 'adbed2c4319687d77322b22c962cf99d', NULL, '2020-10-28 18:33:49', '79.108.8.202', '2020-10-28 18:33:49', 0, 0, 'b1c9b7b7d4c5d0ec'),
(18, 'Subnormal', 'kevinnipero@gmail.com', 'cfd3ec17d607c6254eec2ad3ea9625a0be9135707d87f5675fe1c352f7a1461f', '16ebebb36480491252468ce0a03fac6f', NULL, '2020-10-28 18:34:17', '79.108.8.202', '2020-10-28 18:34:17', 0, 0, '9123ff9bd6b23bee'),
(19, 'LaVuelta', 'balcellsr@hotmail.com', '70709c0d30fe49f68ea65f5565eba34a4aa5f99da3419ed5efe4480e1cfbbd3f', '1dce3ab4744feede71805f294d747df7', NULL, '2020-10-28 18:34:41', '79.108.8.202', '2020-10-28 18:34:41', 0, 0, '9030f7503b083cb6'),
(20, 'Kirito', 'Bchavesetulain@gmail.com', 'be81fd132832d242e4b4d0a7e783c5b06ab5f6e8296d3d0ff68b481a6626abc1', '467defa4c1b0ba882ce7e03067db8b0b', NULL, '2020-10-28 18:35:00', '79.108.8.202', '2020-10-28 18:35:00', 0, 0, '6309f9e0b8455792'),
(21, 'ArminVan', 'santibrocal@hotmail.com', '301f95ade76f0888d001b5f6c48c2d9d3725daf4eb74a84ccc3511ba2442a949', 'b44465fcd80a8011e7d4c0d3c2570a62', NULL, '2020-10-28 18:35:40', '79.108.8.202', '2020-10-28 18:35:40', 0, 0, 'd93f4fd2e9db12de'),
(22, 'Shore', 'ex-.-@live.com.ar', '337534a898a9cd3177c390dad194cf25cb91298ab222203f8affafe57f8a9f91', '998d68cd84e9003c361e8749a0728c27', NULL, '2020-10-28 18:36:20', '79.108.8.202', '2020-10-28 18:36:20', 0, 0, 'd1f813f7b42580e5'),
(23, 'elpadrino', 'acostamati76@gmail.com', '2e7e97b1379c4a6751b881e9d548571ff9a8095ecae4dca5d28f84904faf830d', '8eecb80b3d0d2c4a888d3ddd8e02620f', NULL, '2020-10-28 18:36:41', '79.108.8.202', '2020-10-28 18:36:41', 0, 0, '0e3d8747f7bef956'),
(24, 'Baltica', 'isadonnini9@gmail.com', '629d99fe77c4b07b7e961ae256a5646a4282871143ef3fee99297fa5e60b8563', 'ed37a3ea4d497da665a163a2f32aa23e', NULL, '2020-10-28 18:36:58', '79.108.8.202', '2020-10-28 18:36:58', 0, 0, '1bd7a3ab63c5cbfd');

--
-- Indices de la tabla `cuentas`
--
ALTER TABLE `cuentas`
  ADD PRIMARY KEY (`id`);


--
-- AUTO_INCREMENT de la tabla `cuentas`
--
ALTER TABLE `cuentas`
  MODIFY `id` mediumint(8) UNSIGNED NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=15;