export namespace main {
	
	export class DBConfig {
	    dbType: string;
	    host: string;
	    port: string;
	    database: string;
	    username: string;
	    password: string;
	    tableName: string;
	    connectionType: string;
	    serviceName: string;
	    tnsConnection: string;
	    truncateChars: string;
	
	    static createFrom(source: any = {}) {
	        return new DBConfig(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.dbType = source["dbType"];
	        this.host = source["host"];
	        this.port = source["port"];
	        this.database = source["database"];
	        this.username = source["username"];
	        this.password = source["password"];
	        this.tableName = source["tableName"];
	        this.connectionType = source["connectionType"];
	        this.serviceName = source["serviceName"];
	        this.tnsConnection = source["tnsConnection"];
	        this.truncateChars = source["truncateChars"];
	    }
	}

}

