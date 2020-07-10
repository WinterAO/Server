-- phpMyAdmin SQL Dump
-- version 4.6.6deb5
-- https://www.phpmyadmin.net/
--
-- Servidor: localhost:3306
-- Tiempo de generación: 04-06-2020 a las 19:50:51
-- Versión del servidor: 10.3.22-MariaDB-0+deb10u1
-- Versión de PHP: 7.3.14-1~deb10u1

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
  `salt` varchar(12) NOT NULL,
  `id_recuperacion` varchar(32) DEFAULT NULL,
  `date_created` timestamp NULL DEFAULT current_timestamp(),
  `last_ip` varchar(16) DEFAULT NULL,
  `date_last_login` timestamp NULL DEFAULT current_timestamp(),
  `gemas` int(12) DEFAULT 0,
  `status` tinyint(1) NOT NULL DEFAULT 0,
  `id_confirmacion` varchar(128) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `attribute`
--

CREATE TABLE `attribute` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `number` tinyint(3) UNSIGNED NOT NULL,
  `value` tinyint(3) UNSIGNED NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `bank_item`
--

CREATE TABLE `bank_item` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `number` tinyint(3) UNSIGNED NOT NULL,
  `item_id` smallint(5) UNSIGNED DEFAULT NULL,
  `amount` smallint(5) UNSIGNED DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `inventory_item`
--

CREATE TABLE `inventory_item` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `number` tinyint(3) UNSIGNED NOT NULL,
  `item_id` smallint(5) UNSIGNED DEFAULT NULL,
  `amount` smallint(5) UNSIGNED DEFAULT NULL,
  `is_equipped` tinyint(1) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `pet`
--

CREATE TABLE `pet` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `number` tinyint(3) UNSIGNED NOT NULL,
  `pet_id` smallint(5) UNSIGNED DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

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
-- Estructura de tabla para la tabla `skillpoint`
--

CREATE TABLE `skillpoint` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `number` tinyint(3) UNSIGNED NOT NULL,
  `value` tinyint(3) UNSIGNED NOT NULL,
  `exp` int(10) UNSIGNED NOT NULL,
  `elu` int(10) UNSIGNED NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `spell`
--

CREATE TABLE `spell` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `number` tinyint(3) UNSIGNED NOT NULL,
  `spell_id` smallint(5) UNSIGNED DEFAULT NULL
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
  `guild_rejected_because` varchar(255) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

--
-- Índices para tablas volcadas
--

--
-- Indices de la tabla `account`
--
ALTER TABLE `account`
  ADD PRIMARY KEY (`id`);

--
-- Indices de la tabla `attribute`
--
ALTER TABLE `attribute`
  ADD PRIMARY KEY (`user_id`,`number`);

--
-- Indices de la tabla `bank_item`
--
ALTER TABLE `bank_item`
  ADD PRIMARY KEY (`user_id`,`number`);

--
-- Indices de la tabla `inventory_item`
--
ALTER TABLE `inventory_item`
  ADD PRIMARY KEY (`user_id`,`number`);

--
-- Indices de la tabla `pet`
--
ALTER TABLE `pet`
  ADD PRIMARY KEY (`user_id`,`number`);

--
-- Indices de la tabla `punishment`
--
ALTER TABLE `punishment`
  ADD PRIMARY KEY (`user_id`,`number`);

--
-- Indices de la tabla `skillpoint`
--
ALTER TABLE `skillpoint`
  ADD PRIMARY KEY (`user_id`,`number`);

--
-- Indices de la tabla `spell`
--
ALTER TABLE `spell`
  ADD PRIMARY KEY (`user_id`,`number`);

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
  MODIFY `id` mediumint(8) UNSIGNED NOT NULL AUTO_INCREMENT;
--
-- AUTO_INCREMENT de la tabla `usuario`
--
ALTER TABLE `usuario`
  MODIFY `id` mediumint(8) UNSIGNED NOT NULL AUTO_INCREMENT;
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
