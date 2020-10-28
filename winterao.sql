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
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `npcs1` varchar(64) DEFAULT NULL,
  `estado1` int(1) NOT NULL,
  `npcs2` varchar(64) DEFAULT NULL,
  `estado2` int(1) NOT NULL,
  `npcs3` varchar(64) DEFAULT NULL,
  `estado3` int(1) NOT NULL,
  `npcs4` varchar(64) DEFAULT NULL,
  `estado4` int(1) NOT NULL,
  `npcs5` varchar(64) DEFAULT NULL,
  `estado5` int(1) NOT NULL,
  `npcs6` varchar(64) DEFAULT NULL,
  `estado6` int(1) NOT NULL,
  `npcs7` varchar(64) DEFAULT NULL,
  `estado7` int(1) NOT NULL,
  `npcs8` varchar(64) DEFAULT NULL,
  `estado8` int(1) NOT NULL,
  `npcs9` varchar(64) DEFAULT NULL,
  `estado9` int(1) NOT NULL,
  `npcs10` varchar(64) DEFAULT NULL,
  `estado10` int(1) NOT NULL,
  `npcs11` varchar(64) DEFAULT NULL,
  `estado11` int(1) NOT NULL,
  `npcs12` varchar(64) DEFAULT NULL,
  `estado12` int(1) NOT NULL,
  `npcs13` varchar(64) DEFAULT NULL,
  `estado13` int(1) NOT NULL,
  `npcs14` varchar(64) DEFAULT NULL,
  `estado14` int(1) NOT NULL,
  `npcs15` varchar(64) DEFAULT NULL,
  `estado15` int(1) NOT NULL,
  `npcs16` varchar(64) DEFAULT NULL,
  `estado16` int(1) NOT NULL,
  `npcs17` varchar(64) DEFAULT NULL,
  `estado17` int(1) NOT NULL,
  `npcs18` varchar(64) DEFAULT NULL,
  `estado18` int(1) NOT NULL,
  `npcs19` varchar(64) DEFAULT NULL,
  `estado19` int(1) NOT NULL,
  `npcs20` varchar(64) DEFAULT NULL,
  `estado20` int(1) NOT NULL,
  `npcs21` varchar(64) DEFAULT NULL,
  `estado21` int(1) NOT NULL,
  `npcs22` varchar(64) DEFAULT NULL,
  `estado22` int(1) NOT NULL,
  `npcs23` varchar(64) DEFAULT NULL,
  `estado23` int(1) NOT NULL,
  `npcs24` varchar(64) DEFAULT NULL,
  `estado24` int(1) NOT NULL,
  `npcs25` varchar(64) DEFAULT NULL,
  `estado25` int(1) NOT NULL,
  `npcs26` varchar(64) DEFAULT NULL,
  `estado26` int(1) NOT NULL,
  `npcs27` varchar(64) DEFAULT NULL,
  `estado27` int(1) NOT NULL,
  `npcs28` varchar(64) DEFAULT NULL,
  `estado28` int(1) NOT NULL,
  `npcs29` varchar(64) DEFAULT NULL,
  `estado29` int(1) NOT NULL,
  `npcs30` varchar(64) DEFAULT NULL,
  `estado30` int(1) NOT NULL,
  `npcs31` varchar(64) DEFAULT NULL,
  `estado31` int(1) NOT NULL,
  `npcs32` varchar(64) DEFAULT NULL,
  `estado32` int(1) NOT NULL,
  `npcs33` varchar(64) DEFAULT NULL,
  `estado33` int(1) NOT NULL,
  `npcs34` varchar(64) DEFAULT NULL,
  `estado34` int(1) NOT NULL,
  `npcs35` varchar(64) DEFAULT NULL,
  `estado35` int(1) NOT NULL,
  `npcs36` varchar(64) DEFAULT NULL,
  `estado36` int(1) NOT NULL,
  `npcs37` varchar(64) DEFAULT NULL,
  `estado37` int(1) NOT NULL,
  `npcs38` varchar(64) DEFAULT NULL,
  `estado38` int(1) NOT NULL,
  `npcs39` varchar(64) DEFAULT NULL,
  `estado39` int(1) NOT NULL,
  `npcs40` varchar(64) DEFAULT NULL,
  `estado40` int(1) NOT NULL,
  `npcs41` varchar(64) DEFAULT NULL,
  `estado41` int(1) NOT NULL,
  `npcs42` varchar(64) DEFAULT NULL,
  `estado42` int(1) NOT NULL,
  `npcs43` varchar(64) DEFAULT NULL,
  `estado43` int(1) NOT NULL,
  `npcs44` varchar(64) DEFAULT NULL,
  `estado44` int(1) NOT NULL,
  `npcs45` varchar(64) DEFAULT NULL,
  `estado45` int(1) NOT NULL,
  `npcs46` varchar(64) DEFAULT NULL,
  `estado46` int(1) NOT NULL,
  `npcs47` varchar(64) DEFAULT NULL,
  `estado47` int(1) NOT NULL,
  `npcs48` varchar(64) DEFAULT NULL,
  `estado48` int(1) NOT NULL,
  `npcs49` varchar(64) DEFAULT NULL,
  `estado49` int(1) NOT NULL,
  `npcs50` varchar(64) DEFAULT NULL,
  `estado50` int(1) NOT NULL,
  `npcs51` varchar(64) DEFAULT NULL,
  `estado51` int(1) NOT NULL,
  `npcs52` varchar(64) DEFAULT NULL,
  `estado52` int(1) NOT NULL,
  `npcs53` varchar(64) DEFAULT NULL,
  `estado53` int(1) NOT NULL,
  `npcs54` varchar(64) DEFAULT NULL,
  `estado54` int(1) NOT NULL,
  `npcs55` varchar(64) DEFAULT NULL,
  `estado55` int(1) NOT NULL,
  `npcs56` varchar(64) DEFAULT NULL,
  `estado56` int(1) NOT NULL,
  `npcs57` varchar(64) DEFAULT NULL,
  `estado57` int(1) NOT NULL,
  `npcs58` varchar(64) DEFAULT NULL,
  `estado58` int(1) NOT NULL,
  `npcs59` varchar(64) DEFAULT NULL,
  `estado59` int(1) NOT NULL,
  `npcs60` varchar(64) DEFAULT NULL,
  `estado60` int(1) NOT NULL,
  `npcs61` varchar(64) DEFAULT NULL,
  `estado61` int(1) NOT NULL,
  `npcs62` varchar(64) DEFAULT NULL,
  `estado62` int(1) NOT NULL,
  `npcs63` varchar(64) DEFAULT NULL,
  `estado63` int(1) NOT NULL,
  `npcs64` varchar(64) DEFAULT NULL,
  `estado64` int(1) NOT NULL,
  `npcs65` varchar(64) DEFAULT NULL,
  `estado65` int(1) NOT NULL,
  `npcs66` varchar(64) DEFAULT NULL,
  `estado66` int(1) NOT NULL,
  `npcs67` varchar(64) DEFAULT NULL,
  `estado67` int(1) NOT NULL,
  `npcs68` varchar(64) DEFAULT NULL,
  `estado68` int(1) NOT NULL,
  `npcs69` varchar(64) DEFAULT NULL,
  `estado69` int(1) NOT NULL,
  `npcs70` varchar(64) DEFAULT NULL,
  `estado70` int(1) NOT NULL,
  `npcs71` varchar(64) DEFAULT NULL,
  `estado71` int(1) NOT NULL,
  `npcs72` varchar(64) DEFAULT NULL,
  `estado72` int(1) NOT NULL,
  `npcs73` varchar(64) DEFAULT NULL,
  `estado73` int(1) NOT NULL,
  `npcs74` varchar(64) DEFAULT NULL,
  `estado74` int(1) NOT NULL,
  `npcs75` varchar(64) DEFAULT NULL,
  `estado75` int(1) NOT NULL,
  `npcs76` varchar(64) DEFAULT NULL,
  `estado76` int(1) NOT NULL,
  `npcs77` varchar(64) DEFAULT NULL,
  `estado77` int(1) NOT NULL,
  `npcs78` varchar(64) DEFAULT NULL,
  `estado78`int(1) NOT NULL,
  `npcs79` varchar(64) DEFAULT NULL,
  `estado79` int(1) NOT NULL,
  `npcs80` varchar(64) DEFAULT NULL,
  `estado80` int(1) NOT NULL,
  `npcs81` varchar(64) DEFAULT NULL,
  `estado81` int(1) NOT NULL,
  `npcs82` varchar(64) DEFAULT NULL,
  `estado82` int(1) NOT NULL,
  `npcs83` varchar(64) DEFAULT NULL,
  `estado83` int(1) NOT NULL,
  `npcs84` varchar(64) DEFAULT NULL,
  `estado84` int(1) NOT NULL,
  `npcs85` varchar(64) DEFAULT NULL,
  `estado85` int(1) NOT NULL,
  `npcs86` varchar(64) DEFAULT NULL,
  `estado86` int(1) NOT NULL,
  `npcs87` varchar(64) DEFAULT NULL,
  `estado87` int(1) NOT NULL,
  `npcs88` varchar(64) DEFAULT NULL,
  `estado88` int(1) NOT NULL,
  `npcs89` varchar(64) DEFAULT NULL,
  `estado89` int(1) NOT NULL,
  `npcs90` varchar(64) DEFAULT NULL,
  `estado90` int(1) NOT NULL,
  `npcs91` varchar(64) DEFAULT NULL,
  `estado91` int(1) NOT NULL,
  `npcs92` varchar(64) DEFAULT NULL,
  `estado92` int(1) NOT NULL,
  `npcs93` varchar(64) DEFAULT NULL,
  `estado93` int(1) NOT NULL,
  `npcs94` varchar(64) DEFAULT NULL,
  `estado94` int(1) NOT NULL,
  `npcs95` varchar(64) DEFAULT NULL,
  `estado95` int(1) NOT NULL,
  `npcs96` varchar(64) DEFAULT NULL,
  `estado96` int(1) NOT NULL,
  `npcs97` varchar(64) DEFAULT NULL,
  `estado97` int(1) NOT NULL,
  `npcs98` varchar(64) DEFAULT NULL,
  `estado98` int(1) NOT NULL,
  `npcs99` varchar(64) DEFAULT NULL,
  `estado99` int(1) NOT NULL,
  `npcs100` varchar(64) DEFAULT NULL,
  `estado100` int(1) NOT NULL,
  `npcs101` varchar(64) DEFAULT NULL,
  `estado101` int(1) NOT NULL,
  `npcs102` varchar(64) DEFAULT NULL,
  `estado102` int(1) NOT NULL,
  `npcs103` varchar(64) DEFAULT NULL,
  `estado103` int(1) NOT NULL,
  `npcs104` varchar(64) DEFAULT NULL,
  `estado104` int(1) NOT NULL,
  `npcs105` varchar(64) DEFAULT NULL,
  `estado105` int(1) NOT NULL,
  `npcs106` varchar(64) DEFAULT NULL,
  `estado106` int(1) NOT NULL,
  `npcs107` varchar(64) DEFAULT NULL,
  `estado107` int(1) NOT NULL,
  `npcs108` varchar(64) DEFAULT NULL,
  `estado108` int(1) NOT NULL,
  `npcs109` varchar(64) DEFAULT NULL,
  `estado109` int(1) NOT NULL,
  `npcs110` varchar(64) DEFAULT NULL,
  `estado110` int(1) NOT NULL,
  `npcs111` varchar(64) DEFAULT NULL,
  `estado111` int(1) NOT NULL,
  `npcs112` varchar(64) DEFAULT NULL,
  `estado112` int(1) NOT NULL,
  `npcs113` varchar(64) DEFAULT NULL,
  `estado113` int(1) NOT NULL,
  `npcs114` varchar(64) DEFAULT NULL,
  `estado114` int(1) NOT NULL,
  `npcs115` varchar(64) DEFAULT NULL,
  `estado115` int(1) NOT NULL,
  `npcs116` varchar(64) DEFAULT NULL,
  `estado116` int(1) NOT NULL,
  `npcs117` varchar(64) DEFAULT NULL,
  `estado117` int(1) NOT NULL,
  `npcs118` varchar(64) DEFAULT NULL,
  `estado118` int(1) NOT NULL,
  `npcs119` varchar(64) DEFAULT NULL,
  `estado119` int(1) NOT NULL,
  `npcs120` varchar(64) DEFAULT NULL,
  `estado120` int(1) NOT NULL,
  `npcs121` varchar(64) DEFAULT NULL,
  `estado121` int(1) NOT NULL,
  `npcs122` varchar(64) DEFAULT NULL,
  `estado122` int(1) NOT NULL,
  `npcs123` varchar(64) DEFAULT NULL,
  `estado123` int(1) NOT NULL,
  `npcs124` varchar(64) DEFAULT NULL,
  `estado124` int(1) NOT NULL,
  `npcs125` varchar(64) DEFAULT NULL,
  `estado125` int(1) NOT NULL,
  `npcs126` varchar(64) DEFAULT NULL,
  `estado126` int(1) NOT NULL,
  `npcs127` varchar(64) DEFAULT NULL,
  `estado127` int(1) NOT NULL,
  `npcs128` varchar(64) DEFAULT NULL,
  `estado128` int(1) NOT NULL,
  `npcs129` varchar(64) DEFAULT NULL,
  `estado129` int(1) NOT NULL,
  `npcs130` varchar(64) DEFAULT NULL,
  `estado130` int(1) NOT NULL,
  `npcs131` varchar(64) DEFAULT NULL,
  `estado131` int(1) NOT NULL,
  `npcs132` varchar(64) DEFAULT NULL,
  `estado132` int(1) NOT NULL,
  `npcs133` varchar(64) DEFAULT NULL,
  `estado133` int(1) NOT NULL,
  `npcs134` varchar(64) DEFAULT NULL,
  `estado134` int(1) NOT NULL,
  `npcs135` varchar(64) DEFAULT NULL,
  `estado135` int(1) NOT NULL,
  `npcs136` varchar(64) DEFAULT NULL,
  `estado136` int(1) NOT NULL,
  `npcs137` varchar(64) DEFAULT NULL,
  `estado137` int(1) NOT NULL,
  `npcs138` varchar(64) DEFAULT NULL,
  `estado138` int(1) NOT NULL,
  `npcs139` varchar(64) DEFAULT NULL,
  `estado139` int(1) NOT NULL,
  `npcs140` varchar(64) DEFAULT NULL,
  `estado140` int(1) NOT NULL,
  `npcs141` varchar(64) DEFAULT NULL,
  `estado141` int(1) NOT NULL,
  `npcs142` varchar(64) DEFAULT NULL,
  `estado142` int(1) NOT NULL,
  `npcs143` varchar(64) DEFAULT NULL,
  `estado143` int(1) NOT NULL,
  `npcs144` varchar(64) DEFAULT NULL,
  `estado144` int(1) NOT NULL,
  `npcs145` varchar(64) DEFAULT NULL,
  `estado145` int(1) NOT NULL,
  `npcs146` varchar(64) DEFAULT NULL,
  `estado146` int(1) NOT NULL,
  `npcs147` varchar(64) DEFAULT NULL,
  `estado147` int(1) NOT NULL,
  `npcs148` varchar(64) DEFAULT NULL,
  `estado148` int(1) NOT NULL,
  `npcs149` varchar(64) DEFAULT NULL,
  `estado149` int(1) NOT NULL,
  `npcs150` varchar(64) DEFAULT NULL,
  `estado150` int(1) NOT NULL,
  `npcs151` varchar(64) DEFAULT NULL,
  `estado151` int(1) NOT NULL,
  `npcs152` varchar(64) DEFAULT NULL,
  `estado152` int(1) NOT NULL,
  `npcs153` varchar(64) DEFAULT NULL,
  `estado153` int(1) NOT NULL,
  `npcs154` varchar(64) DEFAULT NULL,
  `estado154` int(1) NOT NULL,
  `npcs155` varchar(64) DEFAULT NULL,
  `estado155` int(1) NOT NULL,
  `npcs156` varchar(64) DEFAULT NULL,
  `estado156` int(1) NOT NULL,
  `npcs157` varchar(64) DEFAULT NULL,
  `estado157` int(1) NOT NULL,
  `npcs158` varchar(64) DEFAULT NULL,
  `estado158` int(1) NOT NULL,
  `npcs159` varchar(64) DEFAULT NULL,
  `estado159` int(1) NOT NULL,
  `npcs160` varchar(64) DEFAULT NULL,
  `estado160` int(1) NOT NULL,
  `npcs161` varchar(64) DEFAULT NULL,
  `estado161` int(1) NOT NULL,
  `npcs162` varchar(64) DEFAULT NULL,
  `estado162` int(1) NOT NULL,
  `npcs163` varchar(64) DEFAULT NULL,
  `estado163` int(1) NOT NULL,
  `npcs164` varchar(64) DEFAULT NULL,
  `estado164` int(1) NOT NULL,
  `npcs165` varchar(64) DEFAULT NULL,
  `estado165` int(1) NOT NULL,
  `npcs166` varchar(64) DEFAULT NULL,
  `estado166` int(1) NOT NULL,
  `npcs167` varchar(64) DEFAULT NULL,
  `estado167` int(1) NOT NULL,
  `npcs168` varchar(64) DEFAULT NULL,
  `estado168` int(1) NOT NULL,
  `npcs169` varchar(64) DEFAULT NULL,
  `estado169` int(1) NOT NULL,
  `npcs170` varchar(64) DEFAULT NULL,
  `estado170` int(1) NOT NULL,
  `npcs171` varchar(64) DEFAULT NULL,
  `estado171` int(1) NOT NULL,
  `npcs172` varchar(64) DEFAULT NULL,
  `estado172` int(1) NOT NULL,
  `npcs173` varchar(64) DEFAULT NULL,
  `estado173` int(1) NOT NULL,
  `npcs174` varchar(64) DEFAULT NULL,
  `estado174` int(1) NOT NULL,
  `npcs175` varchar(64) DEFAULT NULL,
  `estado175` int(1) NOT NULL,
  `npcs176` varchar(64) DEFAULT NULL,
  `estado176` int(1) NOT NULL,
  `npcs177` varchar(64) DEFAULT NULL,
  `estado177` int(1) NOT NULL,
  `npcs178` varchar(64) DEFAULT NULL,
  `estado178` int(1) NOT NULL,
  `npcs179` varchar(64) DEFAULT NULL,
  `estado179` int(1) NOT NULL,
  `npcs180` varchar(64) DEFAULT NULL,
  `estado180` int(1) NOT NULL,
  `npcs181` varchar(64) DEFAULT NULL,
  `estado181` int(1) NOT NULL,
  `npcs182` varchar(64) DEFAULT NULL,
  `estado182` int(1) NOT NULL,
  `npcs183` varchar(64) DEFAULT NULL,
  `estado183` int(1) NOT NULL,
  `npcs184` varchar(64) DEFAULT NULL,
  `estado184` int(1) NOT NULL,
  `npcs185` varchar(64) DEFAULT NULL,
  `estado185` int(1) NOT NULL,
  `npcs186` varchar(64) DEFAULT NULL,
  `estado186` int(1) NOT NULL,
  `npcs187` varchar(64) DEFAULT NULL,
  `estado187` int(1) NOT NULL,
  `npcs188` varchar(64) DEFAULT NULL,
  `estado188` int(1) NOT NULL,
  `npcs189` varchar(64) DEFAULT NULL,
  `estado189` int(1) NOT NULL,
  `npcs190` varchar(64) DEFAULT NULL,
  `estado190` int(1) NOT NULL,
  `npcs191` varchar(64) DEFAULT NULL,
  `estado191` int(1) NOT NULL,
  `npcs192` varchar(64) DEFAULT NULL,
  `estado192` int(1) NOT NULL,
  `npcs193` varchar(64) DEFAULT NULL,
  `estado193` int(1) NOT NULL,
  `npcs194` varchar(64) DEFAULT NULL,
  `estado194` int(1) NOT NULL,
  `npcs195` varchar(64) DEFAULT NULL,
  `estado195` int(1) NOT NULL,
  `npcs196` varchar(64) DEFAULT NULL,
  `estado196` int(1) NOT NULL,
  `npcs197` varchar(64) DEFAULT NULL,
  `estado197` int(1) NOT NULL,
  `npcs198` varchar(64) DEFAULT NULL,
  `estado198` int(1) NOT NULL,
  `npcs199` varchar(64) DEFAULT NULL,
  `estado199` int(1) NOT NULL,
  `npcs200` varchar(64) DEFAULT NULL,
  `estado200` int(1) NOT NULL,
  `npcs201` varchar(64) DEFAULT NULL,
  `estado201` int(1) NOT NULL,
  `npcs202` varchar(64) DEFAULT NULL,
  `estado202` int(1) NOT NULL,
  `npcs203` varchar(64) DEFAULT NULL,
  `estado203` int(1) NOT NULL,
  `npcs204` varchar(64) DEFAULT NULL,
  `estado204` int(1) NOT NULL,
  `npcs205` varchar(64) DEFAULT NULL,
  `estado205` int(1) NOT NULL,
  `npcs206` varchar(64) DEFAULT NULL,
  `estado206` int(1) NOT NULL,
  `npcs207` varchar(64) DEFAULT NULL,
  `estado207` int(1) NOT NULL,
  `npcs208` varchar(64) DEFAULT NULL,
  `estado208` int(1) NOT NULL,
  `npcs209` varchar(64) DEFAULT NULL,
  `estado209` int(1) NOT NULL,
  `npcs210` varchar(64) DEFAULT NULL,
  `estado210` int(1) NOT NULL,
  `npcs211` varchar(64) DEFAULT NULL,
  `estado211` int(1) NOT NULL,
  `npcs212` varchar(64) DEFAULT NULL,
  `estado212` int(1) NOT NULL,
  `npcs213` varchar(64) DEFAULT NULL,
  `estado213` int(1) NOT NULL,
  `npcs214` varchar(64) DEFAULT NULL,
  `estado214` int(1) NOT NULL,
  `npcs215` varchar(64) DEFAULT NULL,
  `estado215` int(1) NOT NULL,
  `npcs216` varchar(64) DEFAULT NULL,
  `estado216` int(1) NOT NULL,
  `npcs217` varchar(64) DEFAULT NULL,
  `estado217` int(1) NOT NULL,
  `npcs218` varchar(64) DEFAULT NULL,
  `estado218` int(1) NOT NULL,
  `npcs219` varchar(64) DEFAULT NULL,
  `estado219` int(1) NOT NULL,
  `npcs220` varchar(64) DEFAULT NULL,
  `estado220` int(1) NOT NULL,
  `npcs221` varchar(64) DEFAULT NULL,
  `estado221` int(1) NOT NULL,
  `npcs222` varchar(64) DEFAULT NULL,
  `estado222` int(1) NOT NULL,
  `npcs223` varchar(64) DEFAULT NULL,
  `estado223` int(1) NOT NULL,
  `npcs224` varchar(64) DEFAULT NULL,
  `estado224` int(1) NOT NULL,
  `npcs225` varchar(64) DEFAULT NULL,
  `estado225` int(1) NOT NULL,
  `npcs226` varchar(64) DEFAULT NULL,
  `estado226` int(1) NOT NULL,
  `npcs227` varchar(64) DEFAULT NULL,
  `estado227` int(1) NOT NULL,
  `npcs228` varchar(64) DEFAULT NULL,
  `estado228` int(1) NOT NULL,
  `npcs229` varchar(64) DEFAULT NULL,
  `estado229` int(1) NOT NULL,
  `npcs230` varchar(64) DEFAULT NULL,
  `estado230` int(1) NOT NULL,
  `npcs231` varchar(64) DEFAULT NULL,
  `estado231` int(1) NOT NULL,
  `npcs232` varchar(64) DEFAULT NULL,
  `estado232` int(1) NOT NULL,
  `npcs233` varchar(64) DEFAULT NULL,
  `estado233` int(1) NOT NULL,
  `npcs234` varchar(64) DEFAULT NULL,
  `estado234` int(1) NOT NULL,
  `npcs235` varchar(64) DEFAULT NULL,
  `estado235` int(1) NOT NULL,
  `npcs236` varchar(64) DEFAULT NULL,
  `estado236` int(1) NOT NULL,
  `npcs237` varchar(64) DEFAULT NULL,
  `estado237` int(1) NOT NULL,
  `npcs238` varchar(64) DEFAULT NULL,
  `estado238` int(1) NOT NULL,
  `npcs239` varchar(64) DEFAULT NULL,
  `estado239` int(1) NOT NULL,
  `npcs240` varchar(64) DEFAULT NULL,
  `estado240` int(1) NOT NULL,
  `npcs241` varchar(64) DEFAULT NULL,
  `estado241` int(1) NOT NULL,
  `npcs242` varchar(64) DEFAULT NULL,
  `estado242` int(1) NOT NULL,
  `npcs243` varchar(64) DEFAULT NULL,
  `estado243` int(1) NOT NULL,
  `npcs244` varchar(64) DEFAULT NULL,
  `estado244` int(1) NOT NULL,
  `npcs245` varchar(64) DEFAULT NULL,
  `estado245` int(1) NOT NULL,
  `npcs246` varchar(64) DEFAULT NULL,
  `estado246` int(1) NOT NULL,
  `npcs247` varchar(64) DEFAULT NULL,
  `estado247` int(1) NOT NULL,
  `npcs248` varchar(64) DEFAULT NULL,
  `estado248` int(1) NOT NULL,
  `npcs249` varchar(64) DEFAULT NULL,
  `estado249` int(1) NOT NULL,
  `npcs250` varchar(64) DEFAULT NULL,
  `estado250` int(1) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `amigos`
--

CREATE TABLE `amigos` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `amigo1` varchar(32) DEFAULT NULL,
  `ignorado1` int(1) NOT NULL,
  `amigo2` varchar(32) DEFAULT NULL,
  `ignorado2` int(1) NOT NULL,
  `amigo3` varchar(32) DEFAULT NULL,
  `ignorado3` int(1) NOT NULL,
  `amigo4` varchar(32) DEFAULT NULL,
  `ignorado4` int(1) NOT NULL,
  `amigo5` varchar(32) DEFAULT NULL,
  `ignorado5` int(1) NOT NULL,
  `amigo6` varchar(32) DEFAULT NULL,
  `ignorado6` int(1) NOT NULL,
  `amigo7` varchar(32) DEFAULT NULL,
  `ignorado7` int(1) NOT NULL,
  `amigo8` varchar(32) DEFAULT NULL,
  `ignorado8` int(1) NOT NULL,
  `amigo9` varchar(32) DEFAULT NULL,
  `ignorado9` int(1) NOT NULL,
  `amigo10` varchar(32) DEFAULT NULL,
  `ignorado10` int(1) NOT NULL,
  `amigo11` varchar(32) DEFAULT NULL,
  `ignorado11` int(1) NOT NULL,
  `amigo12` varchar(32) DEFAULT NULL,
  `ignorado12` int(1) NOT NULL,
  `amigo13` varchar(32) DEFAULT NULL,
  `ignorado13` int(1) NOT NULL,
  `amigo14` varchar(32) DEFAULT NULL,
  `ignorado14` int(1) NOT NULL,
  `amigo15` varchar(32) DEFAULT NULL,
  `ignorado15` int(1) NOT NULL,
  `amigo16` varchar(32) DEFAULT NULL,
  `ignorado16` int(1) NOT NULL,
  `amigo17` varchar(32) DEFAULT NULL,
  `ignorado17` int(1) NOT NULL,
  `amigo18` varchar(32) DEFAULT NULL,
  `ignorado18` int(1) NOT NULL,
  `amigo19` varchar(32) DEFAULT NULL,
  `ignorado19` int(1) NOT NULL,
  `amigo20` varchar(32) DEFAULT NULL,
  `ignorado20` int(1) NOT NULL,
  `amigo21` varchar(32) DEFAULT NULL,
  `ignorado21` int(1) NOT NULL,
  `amigo22` varchar(32) DEFAULT NULL,
  `ignorado22` int(1) NOT NULL,
  `amigo23` varchar(32) DEFAULT NULL,
  `ignorado23` int(1) NOT NULL,
  `amigo24` varchar(32) DEFAULT NULL,
  `ignorado24` int(1) NOT NULL,
  `amigo25` varchar(32) DEFAULT NULL,
  `ignorado25` int(1) NOT NULL,
  `amigo26` varchar(32) DEFAULT NULL,
  `ignorado26` int(1) NOT NULL,
  `amigo27` varchar(32) DEFAULT NULL,
  `ignorado27` int(1) NOT NULL,
  `amigo28` varchar(32) DEFAULT NULL,
  `ignorado28` int(1) NOT NULL,
  `amigo29` varchar(32) DEFAULT NULL,
  `ignorado29` int(1) NOT NULL,
  `amigo30` varchar(32) DEFAULT NULL,
  `ignorado30` int(1) NOT NULL,
  `amigo31` varchar(32) DEFAULT NULL,
  `ignorado31` int(1) NOT NULL,
  `amigo32` varchar(32) DEFAULT NULL,
  `ignorado32` int(1) NOT NULL,
  `amigo33` varchar(32) DEFAULT NULL,
  `ignorado33` int(1) NOT NULL,
  `amigo34` varchar(32) DEFAULT NULL,
  `ignorado34` int(1) NOT NULL,
  `amigo35` varchar(32) DEFAULT NULL,
  `ignorado35` int(1) NOT NULL,
  `amigo36` varchar(32) DEFAULT NULL,
  `ignorado36` int(1) NOT NULL,
  `amigo37` varchar(32) DEFAULT NULL,
  `ignorado37` int(1) NOT NULL,
  `amigo38` varchar(32) DEFAULT NULL,
  `ignorado38` int(1) NOT NULL,
  `amigo39` varchar(32) DEFAULT NULL,
  `ignorado39` int(1) NOT NULL,
  `amigo40` varchar(32) DEFAULT NULL,
  `ignorado40` int(1) NOT NULL,
  `amigo41` varchar(32) DEFAULT NULL,
  `ignorado41` int(1) NOT NULL,
  `amigo42` varchar(32) DEFAULT NULL,
  `ignorado42` int(1) NOT NULL,
  `amigo43` varchar(32) DEFAULT NULL,
  `ignorado43` int(1) NOT NULL,
  `amigo44` varchar(32) DEFAULT NULL,
  `ignorado44` int(1) NOT NULL,
  `amigo45` varchar(32) DEFAULT NULL,
  `ignorado45` int(1) NOT NULL,
  `amigo46` varchar(32) DEFAULT NULL,
  `ignorado46` int(1) NOT NULL,
  `amigo47` varchar(32) DEFAULT NULL,
  `ignorado47` int(1) NOT NULL,
  `amigo48` varchar(32) DEFAULT NULL,
  `ignorado48` int(1) NOT NULL,
  `amigo49` varchar(32) DEFAULT NULL,
  `ignorado49` int(1) NOT NULL,
  `amigo50` varchar(32) DEFAULT NULL,
  `ignorado50` int(1) NOT NULL
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
-- Indices de la tabla `amigos`
--
ALTER TABLE `amigos`
  ADD KEY `fk_amigos_user` (`user_id`);

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
-- Filtros para la tabla `amigos`
--
ALTER TABLE `amigos`
  ADD CONSTRAINT `fk_amigos_user` FOREIGN KEY (`user_id`) REFERENCES `usuario` (`id`);

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
