CREATE DATABASE IF NOT EXISTS luckysheetdb DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
USE luckysheetdb;

CREATE TABLE IF NOT EXISTS `luckysheet` (
  `id` bigint(30) NOT NULL,
  `block_id` varchar(200) NOT NULL,
  `row_col` varchar(50) DEFAULT NULL,
  `index` varchar(200) NOT NULL,
  `list_id` varchar(200) NOT NULL,
  `status` int(1) DEFAULT NULL,
  `json_data` json DEFAULT NULL,
  `order` int(3) DEFAULT NULL,
  `is_delete` int(1) DEFAULT NULL,
  PRIMARY KEY (`id`),
  KEY `lib` (`list_id`,`index`,`block_id`),
  KEY `order_idx` (`order`),
  KEY `status_idx` (`status`),
  KEY `is_delete_idx` (`is_delete`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

INSERT INTO luckysheet VALUES
(139400313311449087, 'fblock', '', '1', '1079500#-8803#7c45f52b7d01486d88bc53cb17dcd2c3', 1, '{"row":84,"name":"Sheet1","chart":[],"color":"","index":"1","order":0,"column":60,"config":{},"status":0,"celldata":[],"ch_width":4748,"rowsplit":[],"rh_height":1790,"scrollTop":0,"scrollLeft":0,"visibledatarow":[],"visibledatacolumn":[],"jfgird_select_save":[],"jfgrid_selection_range":{}}', 0, 0),
(139400313311449088, 'fblock', '', '2', '1079500#-8803#7c45f52b7d01486d88bc53cb17dcd2c3', 0, '{"row":84,"name":"Sheet2","chart":[],"color":"","index":"2","order":1,"column":60,"config":{},"status":0,"celldata":[],"ch_width":4748,"rowsplit":[],"rh_height":1790,"scrollTop":0,"scrollLeft":0,"visibledatarow":[],"visibledatacolumn":[],"jfgird_select_save":[],"jfgrid_selection_range":{}}', 1, 0),
(139400313311449089, 'fblock', '', '3', '1079500#-8803#7c45f52b7d01486d88bc53cb17dcd2c3', 0, '{"row":84,"name":"Sheet3","chart":[],"color":"","index":"3","order":2,"column":60,"config":{},"status":0,"celldata":[],"ch_width":4748,"rowsplit":[],"rh_height":1790,"scrollTop":0,"scrollLeft":0,"visibledatarow":[],"visibledatacolumn":[],"jfgird_select_save":[],"jfgrid_selection_range":{}}', 2, 0);
