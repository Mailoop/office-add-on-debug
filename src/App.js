import React, {Component} from 'react'
import "./App.css"

const DisplayInfo = (params, value) => (
  <div className="info">
    <div className= "info-name" >{params}: </div>
    <div>{value}</div>
    <br></br>
  </div>
)

const DisplayFetchResponse = (name, response) => (
  <div className="info">
    <div className="info-name">{name}</div>
    <div>{response.status}</div>
    <div>{response.ok}</div>
    <br></br>
  </div>
)

class App extends Component {
  constructor(props) {
    super(props)
    this.state = {
      consoleLogSafe: "pending",
      rollbarResponse: {status: "pending" },
      mailoopAppResponse: { status: "pending" },
      exchangeIdtoken: "pending"
    }
  }

  checkRollbar() {
    fetch('https://api.rollbar.com/api/1/status/ping')
      .then(response => this.setState({ rollbarResponse: response }))
      .catch(response => this.setState({ rollbarResponse: { status: "error" } }))

  }

  onExchangeIdSuccess = that => success => {
    that.setState({ exchangeIdtoken: success.value })
  }

  checkGetExchangeIdentityToken() {

    this.props.paramsOffice.mailbox.getUserIdentityTokenAsync(this.onExchangeIdSuccess(this)
    , err => {
        this.setState({ exchangeIdtoken: err.status })
    })
  }

  checkMailoopApp() {
    fetch('https://app.mailoop.com/')
      .then(response => this.setState({ mailoopAppResponse: response }))
      .catch(response => this.setState({ mailoopAppResponse: { status: "error" } }))

  }

  checkConsoleLogSafe() {
    try {
      console.log("Console Log work")
      this.setState({ consoleLogSafe: "Safe"})
    } catch {
      this.setState({ consoleLogSafe: "Not Safe" })
    }
  }

  componentDidMount() {
    this.checkRollbar()
    this.checkMailoopApp()
    this.checkGetExchangeIdentityToken()
    this.checkConsoleLogSafe()
  }

  render() {
    return (
      <div>
        {DisplayInfo("Navigateur", navigator.userAgent )}
        {DisplayInfo("Email", this.props.paramsOffice.email)}
        {DisplayFetchResponse("api.rollbar.com/api/1/status (200 = OK) ",this.state.rollbarResponse)}
        {DisplayInfo("Exchange id Token", this.state.exchangeIdtoken)}
        {DisplayInfo("Console Log Safe ", this.state.consoleLogSafe)}
      </div>
    )
  }
}
  
export default App
