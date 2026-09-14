import app from './index'

const PORT = Number(process.env.PORT || 3001)
app.listen(PORT, () => {
  console.log(`Certificate API running on port ${PORT}`)
})
