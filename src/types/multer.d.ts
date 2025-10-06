declare module 'multer' {
  import type { RequestHandler } from 'express';

  interface MulterOptions {
    storage?: any;
    limits?: {
      fieldNameSize?: number;
      fieldSize?: number;
      fields?: number;
      fileSize?: number;
      files?: number;
      headerPairs?: number;
    };
    preservePath?: boolean;
    fileFilter?: (req: Express.Request, file: any, callback: (error: any, acceptFile: boolean) => void) => void;
  }

  interface MulterFile {
    fieldname: string;
    originalname: string;
    encoding: string;
    mimetype: string;
    size: number;
    destination?: string;
    filename?: string;
    path: string;
    buffer: Buffer;
  }

  interface MulterRequest extends Express.Request {
    file?: MulterFile;
    files?: MulterFile[];
  }

  interface MulterInstance {
    single(field: string): RequestHandler;
    array(field: string, maxCount?: number): RequestHandler;
  }

  function multer(options?: MulterOptions): MulterInstance;
  namespace multer {
    function memoryStorage(): any;
  }

  export = multer;
}


