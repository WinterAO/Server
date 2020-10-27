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
-- Base de datos: `winterao`
--
CREATE DATABASE IF NOT EXISTS `winterao` DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci;
USE `winterao`;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `account`
--

CREATE TABLE `account` (
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
  `id_confirmacion` varchar(128) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

--
-- Volcado de datos para la tabla `account`
--

INSERT INTO `account` (`id`, `username`, `email`, `password`, `salt`, `id_recuperacion`, `date_created`, `last_ip`, `date_last_login`, `gemas`, `status`, `id_confirmacion`) VALUES
(2, 'Lorwik', 'lorwik@gmail.com', 'c7e9f590ade5d41c37e3848339f811a042877bbdc96b02ba5d2461ccfff8127e', 'fDNCB3MyxMGcLpawPkK6GZDv9MADN3i8', NULL, '2020-06-04 18:39:29', '192.168.1.10', '2020-09-05 11:25:44', 0, 1, 'yyUxfFxNP7QySeiUC5P6YK328P5sOh8N'),
(8, 'sANTO', 'sonrisa_eventos@hotmail.com', '71f46a44617e6e6f3f72b68b7980787eec16d69f7d61c52d8ade8b860537764f', 'k0UJiREgZqmXEtZiLfefMEiThYiuSoYJ', NULL, '2020-06-04 19:56:08', '186.138.37.123', '2020-09-03 21:57:25', 0, 1, 'VERIFICADA'),
(10, 'Definiun', 'Definiun@gmail.com', '892a6e79a67699642e8330a24a87d72184842c4dcd097bd9b9f4d346fadb7951', 'D8pgoUVJZYC0vz3imklk6ZUYLXS2gPYh', NULL, '2020-06-04 20:22:58', '192.168.1.7', '2020-09-03 20:45:10', 0, 1, 'JkDnxG7PR0j178DSut7DRHMt4UxSymvY'),
(11, 'jopiodz', 'jopiodz00@hotmail.com', '5e5af03000f398dc00e5d99b796042cfbf7c49fcfe7bc0046c920d43c34d4bff', 'pG79DuVLFKVW6qeoy5xtWudvCk5pCCFP', 'fCxq7Gs4LHZuvjbkckZO4j6Eq0jcEo43', '2020-06-04 20:41:35', '190.244.223.128', '2020-08-30 18:44:34', 0, 1, 'VERIFICADA'),
(12, 'Howell', 'pjs_Del_ao@hotmail.com', '99b6670de91e04a20dd7abf6093beae193f2745bf865f6fd28a2a5cab25c6c6c', '8M2Y7BZ0JtMLWq3DZXgF3JHEJj7ybVmV', NULL, '2020-06-06 00:04:43', '190.50.89.159', '2020-09-03 20:26:04', 0, 1, 'RfwE5NbuCvKbR6Wzduh5krLFFuTmq8zj');

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `attribute`
--

CREATE TABLE `attribute` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `att1` tinyint(3) UNSIGNED NOT NULL,
  `att2` tinyint(3) UNSIGNED NOT NULL,
  `att3` tinyint(3) UNSIGNED NOT NULL,
  `att4` tinyint(3) UNSIGNED NOT NULL,
  `att5` tinyint(3) UNSIGNED NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `bank_item`
--

CREATE TABLE `bank_item` (
      `user_id` mediumint(8) UNSIGNED NOT NULL,
      `item_id1` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount1` smallint(5) UNSIGNED NULL DEFAULT '0',
	  
      `item_id2` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount2` smallint(5) UNSIGNED NULL DEFAULT '0',
    
       `item_id3` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount3` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id4` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount4` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id5` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount5` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id6` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount6` smallint(5) UNSIGNED NULL DEFAULT '0',
   
       `item_id7` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount7` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id8` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount8` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id9` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount9` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id10` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount10` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id11` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount11` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id12` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount12` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id13` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount13` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id14` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount14` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id15` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount15` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id16` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount16` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id17` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount17` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id18` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount18` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id19` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount19` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id20` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount20` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id21` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount21` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id22` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount22` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id23` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount23` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id24` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount24` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id25` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount25` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id26` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount26` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id27` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount27` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id28` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount28` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id29` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount29` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id30` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount30` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id31` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount31` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id32` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount32` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id33` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount33` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id34` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount34` smallint(5) UNSIGNED NULL DEFAULT '0',
	  
	  `item_id35` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount35` smallint(5) UNSIGNED NULL DEFAULT '0',
	  
	  `item_id36` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount36` smallint(5) UNSIGNED NULL DEFAULT '0',
	  
	  `item_id37` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount37` smallint(5) UNSIGNED NULL DEFAULT '0',
	  
	  `item_id38` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount38` smallint(5) UNSIGNED NULL DEFAULT '0',
	  
	  `item_id39` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount39` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id40` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount40` smallint(5) UNSIGNED NULL DEFAULT '0') ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `inventory_item`
--

CREATE TABLE inventory_item (
    `user_id` mediumint(8) UNSIGNED NOT NULL,
      `item_id1` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount1` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped1` tinyint(1) UNSIGNED NULL DEFAULT '0',
	  
      `item_id2` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount2` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped2` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
       `item_id3` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount3` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped3` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id4` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount4` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped4` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id5` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount5` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped5` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id6` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount6` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped6` tinyint(1) UNSIGNED NULL DEFAULT '0',
   
       `item_id7` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount7` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped7` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id8` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount8` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped8` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id9` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount9` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped9` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id10` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount10` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped10` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id11` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount11` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped11` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id12` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount12` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped12` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id13` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount13` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped13` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id14` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount14` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped14` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id15` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount15` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped15` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id16` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount16` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped16` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id17` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount17` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped17` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id18` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount18` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped18` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id19` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount19` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped19` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id20` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount20` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped20` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id21` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount21` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped21` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id22` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount22` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped22` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id23` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount23` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped23` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id24` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount24` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped24` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id25` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount25` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped25` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
       `item_id26` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount26` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped26` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id27` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount27` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped27` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id28` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount28` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped28` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id29` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount29` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped29` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id30` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount30` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped30` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id31` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount31` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped31` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id32` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount32` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped32` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id33` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount33` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped33` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id34` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount34` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped34` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id35` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount35` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped35` tinyint(1) UNSIGNED NULL DEFAULT '0') ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;


--
-- Estructura de tabla para la tabla `pet`
--

CREATE TABLE `pet` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `pet1` smallint(5) UNSIGNED DEFAULT NULL,
  `pet2` smallint(5) UNSIGNED DEFAULT NULL,
  `pet3` smallint(5) UNSIGNED DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `profesion_primaria`
--

CREATE TABLE `profesion_primaria` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `profesion` smallint(4) UNSIGNED NOT NULL DEFAULT 0,
  `receta1` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta2` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta3` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta4` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta5` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta6` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta7` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta8` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta9` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta10` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta11` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta12` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta13` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta14` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta15` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta16` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta17` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta18` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta19` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta20` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta21` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta22` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta23` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta24` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta25` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta26` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta27` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta28` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta29` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta30` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta31` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta32` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta33` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta34` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta35` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta36` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta37` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta38` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta39` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta40` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta41` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta42` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta43` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta44` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta45` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta46` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta47` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta48` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta49` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta50` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta51` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta52` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta53` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta54` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta55` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta56` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta57` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta58` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta59` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta60` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta61` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta62` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta63` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta64` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta65` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta66` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta67` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta68` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta69` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta70` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta71` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta72` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta73` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta74` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta75` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta76` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta77` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta78` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta79` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta80` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta81` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta82` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta83` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta84` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta85` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta86` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta87` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta88` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta89` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta90` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta91` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta92` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta93` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta94` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta95` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta96` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta97` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta98` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta99` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta100` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta101` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta102` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta103` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta104` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta105` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta106` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta107` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta108` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta109` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta110` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta111` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta112` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta113` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta114` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta115` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta116` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta117` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta118` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta119` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta120` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta121` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta122` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta123` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta124` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta125` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta126` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta127` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta128` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta129` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta130` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta131` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta132` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta133` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta134` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta135` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta136` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta137` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta138` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta139` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta140` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta141` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta142` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta143` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta144` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta145` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta146` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta147` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta148` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta149` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta150` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta151` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta152` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta153` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta154` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta155` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta156` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta157` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta158` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta159` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta160` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta161` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta162` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta163` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta164` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta165` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta166` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta167` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta168` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta169` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta170` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta171` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta172` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta173` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta174` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta175` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta176` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta177` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta178` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta179` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta180` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta181` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta182` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta183` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta184` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta185` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta186` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta187` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta188` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta189` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta190` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta191` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta192` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta193` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta194` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta195` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta196` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta197` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta198` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta199` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta200` smallint(6) UNSIGNED NOT NULL DEFAULT 0
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `profesion_secundaria`
--

CREATE TABLE `profesion_secundaria` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `profesion` smallint(4) UNSIGNED NOT NULL DEFAULT 0,
  `receta1` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta2` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta3` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta4` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta5` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta6` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta7` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta8` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta9` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta10` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta11` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta12` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta13` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta14` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta15` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta16` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta17` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta18` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta19` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta20` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta21` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta22` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta23` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta24` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta25` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta26` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta27` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta28` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta29` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta30` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta31` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta32` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta33` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta34` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta35` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta36` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta37` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta38` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta39` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta40` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta41` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta42` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta43` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta44` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta45` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta46` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta47` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta48` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta49` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta50` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta51` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta52` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta53` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta54` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta55` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta56` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta57` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta58` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta59` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta60` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta61` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta62` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta63` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta64` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta65` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta66` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta67` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta68` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta69` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta70` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta71` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta72` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta73` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta74` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta75` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta76` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta77` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta78` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta79` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta80` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta81` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta82` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta83` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta84` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta85` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta86` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta87` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta88` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta89` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta90` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta91` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta92` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta93` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta94` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta95` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta96` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta97` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta98` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta99` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta100` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta101` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta102` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta103` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta104` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta105` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta106` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta107` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta108` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta109` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta110` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta111` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta112` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta113` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta114` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta115` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta116` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta117` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta118` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta119` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta120` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta121` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta122` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta123` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta124` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta125` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta126` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta127` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta128` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta129` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta130` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta131` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta132` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta133` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta134` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta135` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta136` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta137` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta138` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta139` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta140` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta141` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta142` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta143` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta144` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta145` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta146` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta147` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta148` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta149` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta150` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta151` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta152` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta153` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta154` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta155` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta156` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta157` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta158` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta159` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta160` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta161` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta162` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta163` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta164` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta165` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta166` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta167` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta168` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta169` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta170` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta171` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta172` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta173` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta174` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta175` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta176` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta177` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta178` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta179` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta180` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta181` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta182` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta183` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta184` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta185` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta186` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta187` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta188` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta189` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta190` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta191` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta192` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta193` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta194` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta195` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta196` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta197` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta198` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta199` smallint(6) UNSIGNED NOT NULL DEFAULT 0,
  `receta200` smallint(6) UNSIGNED NOT NULL DEFAULT 0
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

--
-- Estructura de tabla para la tabla `punishment`
--

CREATE TABLE `punishment` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `number` tinyint(3) UNSIGNED NOT NULL,
  `reason` varchar(255) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `quest`
--

CREATE TABLE `quest` (
  `idquest` mediumint(8) NOT NULL,
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `completado` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp(),
  `npcs` varchar(256) DEFAULT NULL,
  `estado` int(1) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `skillpoint`
--

CREATE TABLE `skillpoint` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `sk1` tinyint(3) UNSIGNED NOT NULL,
  `exp1` int(10) UNSIGNED NOT NULL,
  `elu1` int(10) UNSIGNED NOT NULL,
  
  `sk2` tinyint(3) UNSIGNED NOT NULL,
  `exp2` int(10) UNSIGNED NOT NULL,
  `elu2` int(10) UNSIGNED NOT NULL,
  
  `sk3` tinyint(3) UNSIGNED NOT NULL,
  `exp3` int(10) UNSIGNED NOT NULL,
  `elu3` int(10) UNSIGNED NOT NULL,
  
  `sk4` tinyint(3) UNSIGNED NOT NULL,
  `exp4` int(10) UNSIGNED NOT NULL,
  `elu4` int(10) UNSIGNED NOT NULL,
  
  `sk5` tinyint(3) UNSIGNED NOT NULL,
  `exp5` int(10) UNSIGNED NOT NULL,
  `elu5` int(10) UNSIGNED NOT NULL,
  
  `sk6` tinyint(3) UNSIGNED NOT NULL,
  `exp6` int(10) UNSIGNED NOT NULL,
  `elu6` int(10) UNSIGNED NOT NULL,
  
  `sk7` tinyint(3) UNSIGNED NOT NULL,
  `exp7` int(10) UNSIGNED NOT NULL,
  `elu7` int(10) UNSIGNED NOT NULL,
  
  `sk8` tinyint(3) UNSIGNED NOT NULL,
  `exp8` int(10) UNSIGNED NOT NULL,
  `elu8` int(10) UNSIGNED NOT NULL,
  
  `sk9` tinyint(3) UNSIGNED NOT NULL,
  `exp9` int(10) UNSIGNED NOT NULL,
  `elu9` int(10) UNSIGNED NOT NULL,
  
  `sk10` tinyint(3) UNSIGNED NOT NULL,
  `exp10` int(10) UNSIGNED NOT NULL,
  `elu10` int(10) UNSIGNED NOT NULL,
  
  `sk11` tinyint(3) UNSIGNED NOT NULL,
  `exp11` int(10) UNSIGNED NOT NULL,
  `elu11` int(10) UNSIGNED NOT NULL,
  
  `sk12` tinyint(3) UNSIGNED NOT NULL,
  `exp12` int(10) UNSIGNED NOT NULL,
  `elu12` int(10) UNSIGNED NOT NULL,
  
  `sk13` tinyint(3) UNSIGNED NOT NULL,
  `exp13` int(10) UNSIGNED NOT NULL,
  `elu13` int(10) UNSIGNED NOT NULL,
  
  `sk14` tinyint(3) UNSIGNED NOT NULL,
  `exp14` int(10) UNSIGNED NOT NULL,
  `elu14` int(10) UNSIGNED NOT NULL,
  
  `sk15` tinyint(3) UNSIGNED NOT NULL,
  `exp15` int(10) UNSIGNED NOT NULL,
  `elu15` int(10) UNSIGNED NOT NULL,
  
  `sk16` tinyint(3) UNSIGNED NOT NULL,
  `exp16` int(10) UNSIGNED NOT NULL,
  `elu16` int(10) UNSIGNED NOT NULL,
  
  `sk17` tinyint(3) UNSIGNED NOT NULL,
  `exp17` int(10) UNSIGNED NOT NULL,
  `elu17` int(10) UNSIGNED NOT NULL,
  
  `sk18` tinyint(3) UNSIGNED NOT NULL,
  `exp18` int(10) UNSIGNED NOT NULL,
  `elu18` int(10) UNSIGNED NOT NULL,
  
  `sk19` tinyint(3) UNSIGNED NOT NULL,
  `exp19` int(10) UNSIGNED NOT NULL,
  `elu19` int(10) UNSIGNED NOT NULL,
  
  `sk20` tinyint(3) UNSIGNED NOT NULL,
  `exp20` int(10) UNSIGNED NOT NULL,
  `elu20` int(10) UNSIGNED NOT NULL,
  
  `sk21` tinyint(3) UNSIGNED NOT NULL,
  `exp21` int(10) UNSIGNED NOT NULL,
  `elu21` int(10) UNSIGNED NOT NULL,
  
  `sk22` tinyint(3) UNSIGNED NOT NULL,
  `exp22` int(10) UNSIGNED NOT NULL,
  `elu22` int(10) UNSIGNED NOT NULL,
  
  `sk23` tinyint(3) UNSIGNED NOT NULL,
  `exp23` int(10) UNSIGNED NOT NULL,
  `elu23` int(10) UNSIGNED NOT NULL,
  
  `sk24` tinyint(3) UNSIGNED NOT NULL,
  `exp24` int(10) UNSIGNED NOT NULL,
  `elu24` int(10) UNSIGNED NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `spell`
--

CREATE TABLE `spell` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `spell_id1` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id2` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id3` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id4` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id5` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id6` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id7` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id8` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id9` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id10` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id11` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id12` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id13` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id14` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id15` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id16` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id17` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id18` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id19` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id20` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id21` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id22` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id23` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id24` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id25` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id26` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id27` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id28` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id29` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id30` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id31` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id32` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id33` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id34` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id35` smallint(5) UNSIGNED DEFAULT 0
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `usuario`
--

CREATE TABLE `usuario` (
  `id` mediumint(8) UNSIGNED NOT NULL,
  `account_id` mediumint(8) UNSIGNED NOT NULL,
  `deleted` tinyint(1) NOT NULL DEFAULT 0,
  `name` varchar(30) NOT NULL,
  `level` smallint(5) UNSIGNED NOT NULL,
  `exp` int(10) UNSIGNED NOT NULL,
  `free_skillpoints` int(10) UNSIGNED NOT NULL,
  `assigned_skillpoints` int(10) UNSIGNED NOT NULL,
  `elu` int(10) UNSIGNED NOT NULL,
  `genre_id` tinyint(3) UNSIGNED NOT NULL,
  `race_id` tinyint(3) UNSIGNED NOT NULL,
  `class_id` tinyint(3) UNSIGNED NOT NULL,
  `home_id` tinyint(3) UNSIGNED NOT NULL,
  `description` varchar(255) DEFAULT NULL,
  `gold` int(10) UNSIGNED NOT NULL,
  `bank_gold` int(10) UNSIGNED NOT NULL DEFAULT 0,
  `elo` int(10) UNSIGNED NOT NULL,
  `pet_amount` tinyint(3) UNSIGNED NOT NULL DEFAULT 0,
  `votes_amount` smallint(5) UNSIGNED DEFAULT 0,
  `pos_map` smallint(5) UNSIGNED NOT NULL,
  `pos_x` tinyint(3) UNSIGNED NOT NULL,
  `pos_y` tinyint(3) UNSIGNED NOT NULL,
  `last_map` tinyint(3) UNSIGNED NOT NULL DEFAULT 1,
  `body_id` smallint(5) UNSIGNED NOT NULL,
  `head_id` smallint(5) UNSIGNED NOT NULL,
  `weapon_id` smallint(5) UNSIGNED NOT NULL,
  `helmet_id` smallint(5) UNSIGNED NOT NULL,
  `shield_id` smallint(5) UNSIGNED NOT NULL,
  `aura_id` int(24) DEFAULT 0,
  `aura_color` int(24) DEFAULT 0,
  `heading` tinyint(3) UNSIGNED NOT NULL DEFAULT 3,
  `items_amount` tinyint(3) UNSIGNED NOT NULL,
  `slot_armour` tinyint(3) UNSIGNED DEFAULT NULL,
  `slot_weapon` tinyint(3) UNSIGNED DEFAULT NULL,
  `slot_helmet` tinyint(3) UNSIGNED DEFAULT NULL,
  `slot_shield` tinyint(3) UNSIGNED DEFAULT NULL,
  `slot_ammo` tinyint(3) UNSIGNED DEFAULT NULL,
  `slot_ship` tinyint(3) UNSIGNED DEFAULT NULL,
  `slot_ring` tinyint(3) UNSIGNED DEFAULT NULL,
  `slot_bag` tinyint(3) UNSIGNED DEFAULT NULL,
  `min_hp` smallint(5) UNSIGNED NOT NULL,
  `max_hp` smallint(5) UNSIGNED NOT NULL,
  `min_man` smallint(5) UNSIGNED NOT NULL,
  `max_man` smallint(5) UNSIGNED NOT NULL,
  `min_sta` smallint(5) UNSIGNED NOT NULL,
  `max_sta` smallint(5) UNSIGNED NOT NULL,
  `min_ham` smallint(5) UNSIGNED NOT NULL,
  `max_ham` smallint(5) UNSIGNED NOT NULL,
  `min_sed` smallint(5) UNSIGNED NOT NULL,
  `max_sed` smallint(5) UNSIGNED NOT NULL,
  `min_hit` smallint(5) UNSIGNED NOT NULL,
  `max_hit` smallint(5) UNSIGNED NOT NULL,
  `killed_npcs` smallint(5) UNSIGNED NOT NULL DEFAULT 0,
  `killed_users` smallint(5) UNSIGNED NOT NULL DEFAULT 0,
  `rep_asesino` mediumint(8) UNSIGNED NOT NULL DEFAULT 0,
  `rep_bandido` mediumint(8) UNSIGNED NOT NULL DEFAULT 0,
  `rep_burgues` mediumint(8) UNSIGNED NOT NULL DEFAULT 0,
  `rep_ladron` mediumint(8) UNSIGNED NOT NULL DEFAULT 0,
  `rep_noble` mediumint(8) UNSIGNED NOT NULL,
  `rep_plebe` mediumint(8) UNSIGNED NOT NULL,
  `rep_average` mediumint(9) NOT NULL,
  `is_naked` tinyint(1) NOT NULL DEFAULT 0,
  `is_poisoned` tinyint(1) NOT NULL DEFAULT 0,
  `is_incinerado` tinyint(1) DEFAULT 0,
  `is_hidden` tinyint(1) NOT NULL DEFAULT 0,
  `is_hungry` tinyint(1) NOT NULL DEFAULT 0,
  `is_thirsty` tinyint(1) NOT NULL DEFAULT 0,
  `is_ban` tinyint(1) NOT NULL DEFAULT 0,
  `is_dead` tinyint(1) NOT NULL DEFAULT 0,
  `is_sailing` tinyint(1) NOT NULL DEFAULT 0,
  `is_paralyzed` tinyint(1) NOT NULL DEFAULT 0,
  `is_logged` tinyint(1) NOT NULL DEFAULT 0,
  `counter_pena` smallint(5) UNSIGNED NOT NULL DEFAULT 0,
  `counter_connected` int(10) UNSIGNED NOT NULL DEFAULT 0,
  `counter_training` int(10) UNSIGNED NOT NULL DEFAULT 0,
  `pertenece_consejo_real` tinyint(1) NOT NULL DEFAULT 0,
  `pertenece_consejo_caos` tinyint(1) NOT NULL DEFAULT 0,
  `pertenece_real` tinyint(1) NOT NULL DEFAULT 0,
  `pertenece_caos` tinyint(1) NOT NULL DEFAULT 0,
  `ciudadanos_matados` smallint(5) UNSIGNED NOT NULL DEFAULT 0,
  `criminales_matados` smallint(5) UNSIGNED NOT NULL DEFAULT 0,
  `recibio_armadura_real` tinyint(1) NOT NULL DEFAULT 0,
  `recibio_armadura_caos` tinyint(1) NOT NULL DEFAULT 0,
  `recibio_exp_real` tinyint(1) NOT NULL DEFAULT 0,
  `recibio_exp_caos` tinyint(1) NOT NULL DEFAULT 0,
  `recompensas_real` tinyint(3) UNSIGNED DEFAULT 0,
  `recompensas_caos` tinyint(3) UNSIGNED DEFAULT 0,
  `reenlistadas` smallint(5) UNSIGNED NOT NULL DEFAULT 0,
  `fecha_ingreso` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp(),
  `nivel_ingreso` smallint(5) UNSIGNED DEFAULT NULL,
  `matados_ingreso` smallint(5) UNSIGNED DEFAULT NULL,
  `siguiente_recompensa` smallint(5) UNSIGNED DEFAULT NULL,
  `guild_index` smallint(5) UNSIGNED DEFAULT 0,
  `guild_aspirant_index` smallint(5) UNSIGNED DEFAULT NULL,
  `guild_member_history` varchar(1024) DEFAULT NULL,
  `guild_requests_history` varchar(1024) DEFAULT NULL,
  `guild_rejected_because` varchar(255) DEFAULT NULL,
  `is_global` tinyint(1) DEFAULT 1,
  `profesionA` int(2) NOT NULL DEFAULT 0,
  `ProfesionB` int(2) NOT NULL DEFAULT 0,
  `modocombate` tinyint(4) DEFAULT 0,
  `seguro` tinyint(1) DEFAULT 0
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

--
-- Indices de la tabla `account`
--
ALTER TABLE `account`
  ADD PRIMARY KEY (`id`);

--
-- Indices de la tabla `attribute`
--
ALTER TABLE `attribute`
  ADD PRIMARY KEY (`user_id`);

--
-- Indices de la tabla `pet`
--
ALTER TABLE `pet`
  ADD PRIMARY KEY (`user_id`);

--
-- Indices de la tabla `punishment`
--
ALTER TABLE `punishment`
  ADD PRIMARY KEY (`user_id`,`number`);

--
-- Indices de la tabla `quest`
--
ALTER TABLE `quest`
  ADD KEY `fk_quest_user` (`user_id`);

--
-- Indices de la tabla `skillpoint`
--
ALTER TABLE `skillpoint`
  ADD PRIMARY KEY (`user_id`);

--
-- Indices de la tabla `spell`
--
ALTER TABLE `spell`
  ADD PRIMARY KEY (`user_id`);

--
-- Indices de la tabla `usuario`
--
ALTER TABLE `usuario`
  ADD PRIMARY KEY (`id`),
  ADD KEY `fk_user_account` (`account_id`),
  ADD KEY `name` (`name`);

--
-- AUTO_INCREMENT de las tablas volcadas
--

--
-- AUTO_INCREMENT de la tabla `account`
--
ALTER TABLE `account`
  MODIFY `id` mediumint(8) UNSIGNED NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=15;
--
-- AUTO_INCREMENT de la tabla `usuario`
--
ALTER TABLE `usuario`
  MODIFY `id` mediumint(8) UNSIGNED NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=15;
--
-- Restricciones para tablas volcadas
--

--
-- Filtros para la tabla `attribute`
--
ALTER TABLE `attribute`
  ADD CONSTRAINT `fk_attribute_user` FOREIGN KEY (`user_id`) REFERENCES `usuario` (`id`);

--
-- Filtros para la tabla `bank_item`
--
ALTER TABLE `bank_item`
  ADD CONSTRAINT `fk_bank_user` FOREIGN KEY (`user_id`) REFERENCES `usuario` (`id`);

--
-- Filtros para la tabla `inventory_item`
--
ALTER TABLE `inventory_item`
  ADD CONSTRAINT `fk_inventory_user` FOREIGN KEY (`user_id`) REFERENCES `usuario` (`id`);

--
-- Filtros para la tabla `pet`
--
ALTER TABLE `pet`
  ADD CONSTRAINT `fk_pet_user` FOREIGN KEY (`user_id`) REFERENCES `usuario` (`id`);

--
-- Filtros para la tabla `punishment`
--
ALTER TABLE `punishment`
  ADD CONSTRAINT `fk_punishment_user` FOREIGN KEY (`user_id`) REFERENCES `usuario` (`id`);

--
-- Filtros para la tabla `quest`
--
ALTER TABLE `quest`
  ADD CONSTRAINT `fk_quest_user` FOREIGN KEY (`user_id`) REFERENCES `usuario` (`id`);

--
-- Filtros para la tabla `skillpoint`
--
ALTER TABLE `skillpoint`
  ADD CONSTRAINT `fk_skillpoint_user` FOREIGN KEY (`user_id`) REFERENCES `usuario` (`id`);

--
-- Filtros para la tabla `spell`
--
ALTER TABLE `spell`
  ADD CONSTRAINT `fk_spell_user` FOREIGN KEY (`user_id`) REFERENCES `usuario` (`id`);

--
-- Filtros para la tabla `usuario`
--
ALTER TABLE `usuario`
  ADD CONSTRAINT `fk_user_account` FOREIGN KEY (`account_id`) REFERENCES `account` (`id`);

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
