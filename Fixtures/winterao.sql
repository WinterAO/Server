-- phpMyAdmin SQL Dump
-- version 5.0.4deb2+deb11u1
-- https://www.phpmyadmin.net/
--
-- Servidor: localhost:3306
-- Tiempo de generación: 16-11-2024 a las 19:26:20
-- Versión del servidor: 10.5.23-MariaDB-0+deb11u1
-- Versión de PHP: 7.4.33

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
START TRANSACTION;
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
-- Estructura de tabla para la tabla `amigos`
--

CREATE TABLE `amigos` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `slot` tinyint(3) UNSIGNED NOT NULL DEFAULT 0,
  `amigo` varchar(32) DEFAULT '',
  `ignorado` int(1) DEFAULT 0
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

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
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `banco_items`
--

CREATE TABLE `banco_items` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `slot` tinyint(3) UNSIGNED NOT NULL,
  `item_id` smallint(5) UNSIGNED DEFAULT NULL,
  `amount` smallint(5) UNSIGNED DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `inventario_items`
--

CREATE TABLE `inventario_items` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `slot` tinyint(3) UNSIGNED NOT NULL,
  `item_id` smallint(5) UNSIGNED DEFAULT NULL,
  `amount` smallint(5) UNSIGNED DEFAULT NULL,
  `is_equipped` tinyint(1) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `macros`
--

CREATE TABLE `macros` (
  `user_id` mediumint(8) UNSIGNED NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `macro_acciones`
--

CREATE TABLE `macro_acciones` (
  `id` int(11) NOT NULL,
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `slot` tinyint(1) UNSIGNED NOT NULL,
  `tipo_accion` tinyint(1) UNSIGNED DEFAULT 0,
  `spell` smallint(5) UNSIGNED DEFAULT 0,
  `inv` smallint(5) UNSIGNED DEFAULT 0,
  `command` varchar(24) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

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
  `seguro` tinyint(1) DEFAULT 0,
  `expPVP` int(10) UNSIGNED NOT NULL,
  `levelPVP` int(10) UNSIGNED NOT NULL,
  `eluPVP` int(10) UNSIGNED NOT NULL,
  `origbody_id` smallint(5) UNSIGNED NOT NULL,
  `orighead_id` smallint(5) UNSIGNED NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `pet`
--

CREATE TABLE `pet` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `slot` tinyint(3) UNSIGNED NOT NULL,
  `pet_id` smallint(5) UNSIGNED DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

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
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

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
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `punishment`
--

CREATE TABLE `punishment` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `number` tinyint(3) UNSIGNED NOT NULL,
  `reason` varchar(255) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `quest`
--

CREATE TABLE `quest` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `quest_id` tinyint(3) UNSIGNED NOT NULL DEFAULT 0,
  `npcs` varchar(64) NOT NULL DEFAULT '',
  `estado` tinyint(1) UNSIGNED NOT NULL DEFAULT 0
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `skillpoint`
--

CREATE TABLE `skillpoint` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `skill_id` tinyint(3) UNSIGNED NOT NULL,
  `sk` tinyint(3) UNSIGNED NOT NULL,
  `exp` int(10) UNSIGNED NOT NULL,
  `elu` int(10) UNSIGNED NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `spell`
--

CREATE TABLE `spell` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `slot` tinyint(3) UNSIGNED NOT NULL,
  `spell_id` smallint(5) UNSIGNED DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Índices para tablas volcadas
--

--
-- Indices de la tabla `amigos`
--
ALTER TABLE `amigos`
  ADD PRIMARY KEY (`user_id`,`slot`);

--
-- Indices de la tabla `atributos`
--
ALTER TABLE `atributos`
  ADD PRIMARY KEY (`user_id`);

--
-- Indices de la tabla `banco_items`
--
ALTER TABLE `banco_items`
  ADD PRIMARY KEY (`user_id`,`slot`);

--
-- Indices de la tabla `inventario_items`
--
ALTER TABLE `inventario_items`
  ADD PRIMARY KEY (`user_id`,`slot`);

--
-- Indices de la tabla `macros`
--
ALTER TABLE `macros`
  ADD PRIMARY KEY (`user_id`);

--
-- Indices de la tabla `macro_acciones`
--
ALTER TABLE `macro_acciones`
  ADD PRIMARY KEY (`id`),
  ADD KEY `user_id` (`user_id`);

--
-- Indices de la tabla `personaje`
--
ALTER TABLE `personaje`
  ADD PRIMARY KEY (`id`),
  ADD KEY `name` (`name`);

--
-- Indices de la tabla `pet`
--
ALTER TABLE `pet`
  ADD PRIMARY KEY (`user_id`,`slot`);

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
-- Indices de la tabla `skillpoint`
--
ALTER TABLE `skillpoint`
  ADD PRIMARY KEY (`user_id`,`skill_id`);

--
-- Indices de la tabla `spell`
--
ALTER TABLE `spell`
  ADD PRIMARY KEY (`user_id`,`slot`);

--
-- AUTO_INCREMENT de las tablas volcadas
--

--
-- AUTO_INCREMENT de la tabla `macro_acciones`
--
ALTER TABLE `macro_acciones`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT;

--
-- AUTO_INCREMENT de la tabla `personaje`
--
ALTER TABLE `personaje`
  MODIFY `id` mediumint(8) UNSIGNED NOT NULL AUTO_INCREMENT;

--
-- Restricciones para tablas volcadas
--

--
-- Filtros para la tabla `amigos`
--
ALTER TABLE `amigos`
  ADD CONSTRAINT `fk_amigos_user` FOREIGN KEY (`user_id`) REFERENCES `personaje` (`id`);

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
-- Filtros para la tabla `macro_acciones`
--
ALTER TABLE `macro_acciones`
  ADD CONSTRAINT `macro_acciones_ibfk_1` FOREIGN KEY (`user_id`) REFERENCES `macros` (`user_id`);

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
-- Filtros para la tabla `skillpoint`
--
ALTER TABLE `skillpoint`
  ADD CONSTRAINT `fk_skillpoint_user` FOREIGN KEY (`user_id`) REFERENCES `personaje` (`id`);

--
-- Filtros para la tabla `spell`
--
ALTER TABLE `spell`
  ADD CONSTRAINT `fk_spell_user` FOREIGN KEY (`user_id`) REFERENCES `personaje` (`id`);
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
