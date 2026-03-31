module.exports = {
  run: [
    {
      method: "shell.run",
      params: {
        path: "app",
        message: [
          'npm install --no-audit --no-fund'
        ]
      }
    }
  ]
}
