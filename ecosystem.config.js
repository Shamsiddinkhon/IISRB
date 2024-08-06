module.exports = {
   apps: [
      {
         name: "grafik", // Replace with your application's name
         script: "grafik.js", // Replace with the entry point of your application
         watch: true,
         autorestart: true,
         restart_delay: 60000, // Delay before restarting (in milliseconds), here it's set to 1 minute
         max_restarts: 100, // Optional: limit the number of restarts
         env: {
            NODE_ENV: "development",
         },
         env_production: {
            NODE_ENV: "production",
         },
      },
   ],
};
