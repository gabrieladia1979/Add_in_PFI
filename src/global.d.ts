declare module "*.css";

// Variables inyectadas por Webpack DefinePlugin en tiempo de compilación
declare const process: {
  env: {
    API_BASE_URL?: string;
    API_KEY?: string;
  };
};
