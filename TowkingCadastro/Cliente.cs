using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text;

namespace TowkingCadastro
{
    public class Cliente : INotifyPropertyChanged
    {
        private string _nome;
        public string Nome
        {
            get { return this._nome; }
            set 
            {
                if (value != this._nome)
                    this._nome = value;
                NotifyPropertyChanged();
            }
        }
        private string _sobrenome;
        public string Sobrenome
        {
            get { return this._sobrenome; }
            set
            {
                if (value != this._sobrenome)
                    this._sobrenome = value;
                NotifyPropertyChanged();
            }
        }
        private string _nascimento;
        public string Nascimento
        {
            get { return this._nascimento; }
            set
            {
                if (value != this._nascimento)
                    this._nascimento = value;
                NotifyPropertyChanged();
            }
        }
        private string _telefone;
        public string Telefone
        {
            get { return this._telefone; }
            set
            {
                if (value != this._telefone)
                    this._telefone = value;
                NotifyPropertyChanged();
            }
        }
        private string _celular;
        public string Celular
        {
            get { return this._celular; }
            set
            {
                if (value != this._celular)
                    this._celular = value;
                NotifyPropertyChanged();
            }
        }
        private string _cpf;
        public string CPF
        {
            get { return this._cpf; }
            set
            {
                if (value != this._cpf)
                    this._cpf = value;
                NotifyPropertyChanged();
            }
        }
        private string _email;
        public string Email
        {
            get { return this._email; }
            set
            {
                if (value != this._email)
                    this._email = value;
                NotifyPropertyChanged();
            }
        }
        private string _tipoVeiculo;
        public string TipoVeiculo
        {
            get { return this._tipoVeiculo; }
            set
            {
                if (value != this._tipoVeiculo)
                    this._tipoVeiculo = value;
                NotifyPropertyChanged();
            }
        }
        private string _modeloVeiculo;
        public string ModeloVeiculo
        {
            get { return this._modeloVeiculo; }
            set
            {
                if (value != this._modeloVeiculo)
                    this._modeloVeiculo = value;
                NotifyPropertyChanged();
            }
        }
        private string _rua;
        public string Rua
        {
            get { return this._rua; }
            set
            {
                if (value != this._rua)
                    this._rua = value;
                NotifyPropertyChanged();
            }
        }
        private string _bairro;
        public string Bairro
        {
            get { return this._bairro; }
            set
            {
                if (value != this._bairro)
                    this._bairro = value;
                NotifyPropertyChanged();
            }
        }
        private string _cidade;
        public string Cidade
        {
            get { return this._cidade; }
            set
            {
                if (value != this._cidade)
                    this._cidade = value;
                NotifyPropertyChanged();
            }
        }
        private string _estado;
        public string Estado
        {
            get { return this._estado; }
            set
            {
                if (value != this._estado)
                    this._estado = value;
                NotifyPropertyChanged();
            }
        }
        private string _cep;
        public string CEP
        {
            get { return this._cep; }
            set
            {
                if (value != this._cep)
                    this._cep = value;
                NotifyPropertyChanged();
            }
        }

        private bool _isempresa;
        public bool isempresa
        {
            get { return this._isempresa; }
            set
            {
                if (value != this._isempresa)
                    this._isempresa = value;
                NotifyPropertyChanged();
            }
        }

        public Cliente()
        {

        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

    }
}
