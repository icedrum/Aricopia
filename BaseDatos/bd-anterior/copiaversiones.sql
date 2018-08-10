/*
SQLyog - Free MySQL GUI v5.18
Host - 5.0.45-community-nt : Database - copiaversiones
*********************************************************************
Server version : 5.0.45-community-nt
*/

SET NAMES utf8;

SET SQL_MODE='';

SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0;
SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO';

/*Table structure for table `aplicaciones` */

CREATE TABLE `aplicaciones` (
  `idaplicacion` int(4) NOT NULL,
  `nombre` varchar(40) NOT NULL,
  `ultimaversion` varchar(10) NOT NULL,
  `servidor` varchar(100) NOT NULL,
  `pathservidor` varchar(100) NOT NULL,
  `ejecutable` varchar(40) NOT NULL,
  PRIMARY KEY  (`idaplicacion`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Table structure for table `ficheroscopia` */

CREATE TABLE `ficheroscopia` (
  `idfic` int(4) NOT NULL,
  `idaplicacion` int(4) NOT NULL,
  `nombre` varchar(100) default NULL,
  `tipo` tinyint(1) NOT NULL COMMENT '0=fichero 1=carpeta',
  PRIMARY KEY  (`idfic`,`idaplicacion`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Table structure for table `pcscopia` */

CREATE TABLE `pcscopia` (
  `idpcs` int(4) NOT NULL,
  `idaplicacion` int(4) NOT NULL,
  `nompc` varchar(100) NOT NULL,
  `descripcion` varchar(100) default NULL,
  `pathcopia` varchar(100) NOT NULL,
  PRIMARY KEY  (`idpcs`,`idaplicacion`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

SET SQL_MODE=@OLD_SQL_MODE;
SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS;
