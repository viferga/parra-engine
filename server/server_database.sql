-- phpMyAdmin SQL Dump
-- version 2.9.0.2
-- http://www.phpmyadmin.net
-- 
-- Servidor: localhost
-- Tiempo de generación: 31-12-2009 a las 02:49:53
-- Versión del servidor: 5.0.24
-- Versión de PHP: 4.4.4
-- 
-- Base de datos: `server_database`
-- 

-- --------------------------------------------------------

-- 
-- Estructura de tabla para la tabla `cuentas`
-- 

CREATE TABLE `cuentas` (
  `ID_cuenta` smallint(6) NOT NULL auto_increment,
  `accountname` varchar(32) default NULL,
  `password` varchar(64) default NULL,
  `email` varchar(64) default NULL,
  `ban` varchar(1) default NULL,
  `pj1` varchar(32) default NULL,
  `pj2` varchar(32) default NULL,
  `pj3` varchar(32) default NULL,
  `pj4` varchar(32) default NULL,
  `pj5` varchar(32) default NULL,
  `pj6` varchar(32) default NULL,
  PRIMARY KEY  (`ID_cuenta`)
) ENGINE=MyISAM  DEFAULT CHARSET=latin1 ROW_FORMAT=DYNAMIC COMMENT='Cuentas MMORPG' AUTO_INCREMENT=19 ;

-- 
-- Volcar la base de datos para la tabla `cuentas`
-- 

INSERT INTO `cuentas` VALUES (1, 'Parra', 'asd', 'parra@hotmail.com', '0', 'Parra', NULL, NULL, NULL, NULL, NULL);
INSERT INTO `cuentas` VALUES (2, 'test', 'a', 'a@a.com', '0', 'NULL', 'NULL', 'NULL', 'NULL', 'NULL', 'NULL');
INSERT INTO `cuentas` VALUES (3, 'test2id', 'asdd', 'aa@a.com', '0', 'NULL', 'NULL', 'NULL', 'NULL', 'NULL', 'NULL');
INSERT INTO `cuentas` VALUES (4, 'iasdd12', 'aaaa', 'asmd@adsd.com', '0', 'NULL', 'NULL', 'NULL', 'NULL', 'NULL', 'NULL');
INSERT INTO `cuentas` VALUES (5, 'eeaass', 'asssdf', 'ass@asje.com', '0', 'NULL', 'NULL', 'NULL', 'NULL', 'NULL', 'NULL');
INSERT INTO `cuentas` VALUES (6, 'asjks', 'assskf', 'aas@aksjd.com', '0', 'NULL', 'NULL', 'NULL', 'NULL', 'NULL', 'NULL');
INSERT INTO `cuentas` VALUES (7, 'iiiss', 'sssa', 'aaas@a.com', '0', 'NULL', 'NULL', 'NULL', 'NULL', 'NULL', 'NULL');
INSERT INTO `cuentas` VALUES (8, 'a', 'a', 'a@a.com', '0', 'NULL', 'NULL', 'NULL', 'NULL', 'NULL', 'NULL');
INSERT INTO `cuentas` VALUES (9, 'nk', 'nk34carajo', 'nick-dead@hotmail.com', '0', 'NULL', 'NULL', 'NULL', 'NULL', 'NULL', 'NULL');
INSERT INTO `cuentas` VALUES (10, 'cuenttt', 'asd', 'a@e.com', '0', 'NULL', 'NULL', 'NULL', 'NULL', 'NULL', 'NULL');
INSERT INTO `cuentas` VALUES (13, 'aggrr', 'a', 'a', '0', 'NULL', 'NULL', 'NULL', 'NULL', 'NULL', 'NULL');
INSERT INTO `cuentas` VALUES (14, 'adfddKK', 'aaa', 'aaa', '0', 'NULL', 'NULL', 'NULL', 'NULL', 'NULL', 'NULL');
INSERT INTO `cuentas` VALUES (15, 'errr', 'asd', 'asd', '0', 'NULL', 'NULL', 'NULL', 'NULL', 'NULL', 'NULL');
INSERT INTO `cuentas` VALUES (16, 'adsdd', 'ult', 'ult', '0', 'NULL', 'NULL', 'NULL', 'NULL', 'NULL', 'NULL');
INSERT INTO `cuentas` VALUES (18, 'Jose', 'chamot', 'chamot11@hotmail.com', '0', 'NULL', 'NULL', 'NULL', 'NULL', 'NULL', 'NULL');

-- --------------------------------------------------------

-- 
-- Estructura de tabla para la tabla `players`
-- 

CREATE TABLE `players` (
  `playername` varchar(32) default NULL,
  `ban` int(1) default NULL,
  `lastip` varchar(15) default NULL,
  `body` int(3) default NULL,
  `head` int(3) default NULL,
  `heading` int(1) default NULL,
  `map` int(3) default NULL,
  `x` int(3) default NULL,
  `y` int(3) default NULL
) ENGINE=MyISAM DEFAULT CHARSET=utf8;

-- 
-- Volcar la base de datos para la tabla `players`
-- 

INSERT INTO `players` VALUES ('Parra', 0, '127.0.0.1', 1, 1, 1, 1, 50, 50);
