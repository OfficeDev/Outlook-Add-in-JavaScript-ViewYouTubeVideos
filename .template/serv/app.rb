require 'sinatra/base'
require 'webrick'
require 'webrick/https'
require 'openssl'

CERTS = 'cert/'

opts = {
  :Port => <?bash $port ?>,
  :DocumentRoot => "../app/youtube.html",
  :Logger => WEBrick::Log::new($stderr, WEBrick::Log::DEBUG),
  :DoNotReverseLookup => true,
  :SSLEnable => true,
  :SSLVerifyClient => OpenSSL::SSL::VERIFY_NONE,
  :SSLCertificate => OpenSSL::X509::Certificate.new(File.open(File.join(CERTS, "server.crt")).read),
  :SSLPrivateKey => OpenSSL::PKey::RSA.new(File.open(File.join(CERTS, "server.key")).read),
  :SSLCertName => [ [ "CN",WEBrick::Utils::getservername ] ]
}

class AddInServer < Sinatra::Base
    # Define any custom paths here.
end

Rack::Handler::WEBrick.run AddInServer, opts
