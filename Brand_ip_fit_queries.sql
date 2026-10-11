SELECT * FROM sponsoriq.brand_ip_fit;
SET FOREIGN_KEY_CHECKS = 0;

TRUNCATE TABLE brand_ip_fit;

SET FOREIGN_KEY_CHECKS = 1;
alter table brand_ip_fit drop column Recommendation_Reason;