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
-- Estructura de tabla para la tabla `atributos`
--

CREATE TABLE `atributos` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `att1` tinyint(3) UNSIGNED NOT NULL,
  `att2` tinyint(3) UNSIGNED NOT NULL,
  `att3` tinyint(3) UNSIGNED NOT NULL,
  `att4` tinyint(3) UNSIGNED NOT NULL,
  `att5` tinyint(3) UNSIGNED NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `banco_items`
--

CREATE TABLE `banco_items` (
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
-- Estructura de tabla para la tabla `inventario_items`
--

CREATE TABLE inventario_items (
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
  `quest_id` tinyint(3) UNSIGNED NOT NULL DEFAULT 0,
  `npcs` varchar(64) NOT NULL DEFAULT '',
  `estado` tinyint(1) UNSIGNED NOT NULL DEFAULT 0
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `amigos`
--

CREATE TABLE `amigos` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `amigo1` varchar(32) DEFAULT '',
  `ignorado1` INT(1) NULL DEFAULT '0',
  `amigo2` varchar(32) DEFAULT '',
  `ignorado2` INT(1) NULL DEFAULT '0',
  `amigo3` varchar(32) DEFAULT '',
  `ignorado3` INT(1) NULL DEFAULT '0',
  `amigo4` varchar(32) DEFAULT '',
  `ignorado4` INT(1) NULL DEFAULT '0',
  `amigo5` varchar(32) DEFAULT '',
  `ignorado5` INT(1) NULL DEFAULT '0',
  `amigo6` varchar(32) DEFAULT '',
  `ignorado6` INT(1) NULL DEFAULT '0',
  `amigo7` varchar(32) DEFAULT '',
  `ignorado7` INT(1) NULL DEFAULT '0',
  `amigo8` varchar(32) DEFAULT '',
  `ignorado8` INT(1) NULL DEFAULT '0',
  `amigo9` varchar(32) DEFAULT '',
  `ignorado9` INT(1) NULL DEFAULT '0',
  `amigo10` varchar(32) DEFAULT '',
  `ignorado10` INT(1) NULL DEFAULT '0',
  `amigo11` varchar(32) DEFAULT '',
  `ignorado11` INT(1) NULL DEFAULT '0',
  `amigo12` varchar(32) DEFAULT '',
  `ignorado12` INT(1) NULL DEFAULT '0',
  `amigo13` varchar(32) DEFAULT '',
  `ignorado13` INT(1) NULL DEFAULT '0',
  `amigo14` varchar(32) DEFAULT '',
  `ignorado14` INT(1) NULL DEFAULT '0',
  `amigo15` varchar(32) DEFAULT '',
  `ignorado15` INT(1) NULL DEFAULT '0',
  `amigo16` varchar(32) DEFAULT '',
  `ignorado16` INT(1) NULL DEFAULT '0',
  `amigo17` varchar(32) DEFAULT '',
  `ignorado17` INT(1) NULL DEFAULT '0',
  `amigo18` varchar(32) DEFAULT '',
  `ignorado18` INT(1) NULL DEFAULT '0',
  `amigo19` varchar(32) DEFAULT '',
  `ignorado19` INT(1) NULL DEFAULT '0',
  `amigo20` varchar(32) DEFAULT '',
  `ignorado20` INT(1) NULL DEFAULT '0',
  `amigo21` varchar(32) DEFAULT '',
  `ignorado21` INT(1) NULL DEFAULT '0',
  `amigo22` varchar(32) DEFAULT '',
  `ignorado22` INT(1) NULL DEFAULT '0',
  `amigo23` varchar(32) DEFAULT '',
  `ignorado23` INT(1) NULL DEFAULT '0',
  `amigo24` varchar(32) DEFAULT '',
  `ignorado24` INT(1) NULL DEFAULT '0',
  `amigo25` varchar(32) DEFAULT '',
  `ignorado25` INT(1) NULL DEFAULT '0',
  `amigo26` varchar(32) DEFAULT '',
  `ignorado26` INT(1) NULL DEFAULT '0',
  `amigo27` varchar(32) DEFAULT '',
  `ignorado27` INT(1) NULL DEFAULT '0',
  `amigo28` varchar(32) DEFAULT '',
  `ignorado28` INT(1) NULL DEFAULT '0',
  `amigo29` varchar(32) DEFAULT '',
  `ignorado29` INT(1) NULL DEFAULT '0',
  `amigo30` varchar(32) DEFAULT '',
  `ignorado30` INT(1) NULL DEFAULT '0',
  `amigo31` varchar(32) DEFAULT '',
  `ignorado31` INT(1) NULL DEFAULT '0',
  `amigo32` varchar(32) DEFAULT '',
  `ignorado32` INT(1) NULL DEFAULT '0',
  `amigo33` varchar(32) DEFAULT '',
  `ignorado33` INT(1) NULL DEFAULT '0',
  `amigo34` varchar(32) DEFAULT '',
  `ignorado34` INT(1) NULL DEFAULT '0',
  `amigo35` varchar(32) DEFAULT '',
  `ignorado35` INT(1) NULL DEFAULT '0',
  `amigo36` varchar(32) DEFAULT '',
  `ignorado36` INT(1) NULL DEFAULT '0',
  `amigo37` varchar(32) DEFAULT '',
  `ignorado37` INT(1) NULL DEFAULT '0',
  `amigo38` varchar(32) DEFAULT '',
  `ignorado38` INT(1) NULL DEFAULT '0',
  `amigo39` varchar(32) DEFAULT '',
  `ignorado39` INT(1) NULL DEFAULT '0',
  `amigo40` varchar(32) DEFAULT '',
  `ignorado40` INT(1) NULL DEFAULT '0',
  `amigo41` varchar(32) DEFAULT '',
  `ignorado41` INT(1) NULL DEFAULT '0',
  `amigo42` varchar(32) DEFAULT '',
  `ignorado42` INT(1) NULL DEFAULT '0',
  `amigo43` varchar(32) DEFAULT '',
  `ignorado43` INT(1) NULL DEFAULT '0',
  `amigo44` varchar(32) DEFAULT '',
  `ignorado44` INT(1) NULL DEFAULT '0',
  `amigo45` varchar(32) DEFAULT '',
  `ignorado45` INT(1) NULL DEFAULT '0',
  `amigo46` varchar(32) DEFAULT '',
  `ignorado46` INT(1) NULL DEFAULT '0',
  `amigo47` varchar(32) DEFAULT '',
  `ignorado47` INT(1) NULL DEFAULT '0',
  `amigo48` varchar(32) DEFAULT '',
  `ignorado48` INT(1) NULL DEFAULT '0',
  `amigo49` varchar(32) DEFAULT '',
  `ignorado49` INT(1) NULL DEFAULT '0',
  `amigo50` varchar(32) DEFAULT '',
  `ignorado50` INT(1) NULL DEFAULT '0'
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
-- Estructura de tabla para la tabla `personaje`
--

CREATE TABLE `personaje` (
  `id` mediumint(8) UNSIGNED NOT NULL,
  `cuenta_id` mediumint(8) UNSIGNED NOT NULL,
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
  `pos_map` smallint(3) UNSIGNED NOT NULL,
  `pos_x` int(4) UNSIGNED NOT NULL,
  `pos_y` int(4) UNSIGNED NOT NULL,
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
-- Indices de la tabla `atributos`
--
ALTER TABLE `atributos`
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
  ADD PRIMARY KEY (`user_id`,`quest_id`);
  
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
-- Indices de la tabla `personaje`
--
ALTER TABLE `personaje`
  ADD PRIMARY KEY (`id`),
  ADD KEY `name` (`name`);

--
-- AUTO_INCREMENT de las tablas volcadas
--

--
-- AUTO_INCREMENT de la tabla `personaje`
--
ALTER TABLE `personaje`
  MODIFY `id` mediumint(8) UNSIGNED NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=15;
--
-- Restricciones para tablas volcadas
--

--
-- Filtros para la tabla `atributos`
--
ALTER TABLE `atributos`
  ADD CONSTRAINT `fk_atributos_user` FOREIGN KEY (`user_id`) REFERENCES `personaje` (`id`);

--
-- Filtros para la tabla `banco_items`
--
ALTER TABLE `banco_items`
  ADD CONSTRAINT `fk_bank_user` FOREIGN KEY (`user_id`) REFERENCES `personaje` (`id`);

--
-- Filtros para la tabla `inventario_items`
--
ALTER TABLE `inventario_items`
  ADD CONSTRAINT `fk_inventory_user` FOREIGN KEY (`user_id`) REFERENCES `personaje` (`id`);

--
-- Filtros para la tabla `pet`
--
ALTER TABLE `pet`
  ADD CONSTRAINT `fk_pet_user` FOREIGN KEY (`user_id`) REFERENCES `personaje` (`id`);

--
-- Filtros para la tabla `punishment`
--
ALTER TABLE `punishment`
  ADD CONSTRAINT `fk_punishment_user` FOREIGN KEY (`user_id`) REFERENCES `personaje` (`id`);

--
-- Filtros para la tabla `quest`
--
ALTER TABLE `quest`
  ADD CONSTRAINT `fk_quest_user` FOREIGN KEY (`user_id`) REFERENCES `personaje` (`id`);
  
--
-- Filtros para la tabla `amigos`
--
ALTER TABLE `amigos`
  ADD CONSTRAINT `fk_amigos_user` FOREIGN KEY (`user_id`) REFERENCES `personaje` (`id`);

--
-- Filtros para la tabla `skillpoint`
--
ALTER TABLE `skillpoint`
  ADD CONSTRAINT `fk_skillpoint_user` FOREIGN KEY (`user_id`) REFERENCES `personaje` (`id`);

--
-- Filtros para la tabla `spell`
--
ALTER TABLE `spell`
  ADD CONSTRAINT `fk_spell_user` FOREIGN KEY (`user_id`) REFERENCES `personaje` (`id`);

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;